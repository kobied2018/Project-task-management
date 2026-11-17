import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { after, before, test } from 'node:test';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
let directory;
let server;
let baseUrl;

before(async () => {
  directory = await mkdtemp(join(tmpdir(), 'fieldnotes-follow-ups-'));
  server = spawn(process.execPath, [join(root, 'index.mjs')], {
    env: { ...process.env, NODE_ENV: 'development', PORT: '0', DB_PATH: join(directory, 'test.db') },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  await new Promise((resolve, reject) => {
    let output = '';
    const timer = setTimeout(() => reject(new Error(`Test server did not start: ${output}`)), 15000);
    server.on('error', (error) => { clearTimeout(timer); reject(error); });
    server.on('exit', (code) => { clearTimeout(timer); reject(new Error(`Test server exited (${code}): ${output}`)); });
    server.stderr.on('data', (data) => { output += data; });
    server.stdout.on('data', (data) => {
      output += data;
      const match = output.match(/Fieldnotes API listening on (\d+)/);
      if (match) {
        baseUrl = `http://127.0.0.1:${match[1]}`;
        clearTimeout(timer);
        resolve();
      }
    });
  });
});

after(async () => {
  if (server && server.exitCode === null) {
    const exited = once(server, 'exit');
    server.kill();
    await exited;
  }
  if (directory) await rm(directory, { recursive: true, force: true });
});

async function request(path, { method = 'GET', body, user = 'local-alex', status = 200 } = {}) {
  const response = await fetch(`${baseUrl}/api/projects/northstar${path}`, {
    method,
    headers: { 'Content-Type': 'application/json', 'x-dev-user': user },
    ...(body === undefined ? {} : { body: JSON.stringify(body) }),
  });
  const data = response.status === 204 ? undefined : await response.json();
  assert.equal(response.status, status, JSON.stringify(data));
  return data;
}

async function createTask(title, extra = {}, user = 'local-alex') {
  return (await request('/tasks', {
    method: 'POST', status: 201, user,
    body: { title, topicId: 'launch', status: 'Not started', ...extra },
  })).task;
}

test('follow-ups persist with independent fields, comments, and permissions', async () => {
  const parent = await createTask('Completed parent', { status: 'Done' });
  const child = await createTask('Follow up after completion', {
    parentTaskId: parent.id, owner: 'Test owner', dueDate: '2026-12-01',
  }, 'local-jordan');
  await request(`/tasks/${child.id}`, {
    method: 'PATCH', user: 'local-jordan', body: { status: 'In progress', comments: ['Follow-up progress'] },
  });
  const { project } = await request('');
  assert.equal(project.tasks.find((task) => task.id === parent.id).status, 'Done');
  assert.deepEqual(project.tasks.find((task) => task.id === child.id), {
    ...child, status: 'In progress', comments: ['Follow-up progress'],
  });
  const db = new DatabaseSync(join(directory, 'test.db'), { readOnly: true });
  try {
    const saved = JSON.parse(db.prepare('SELECT state FROM projects WHERE id = ?').get('northstar').state);
    assert.equal(saved.tasks.find((task) => task.id === child.id).parentTaskId, parent.id);
  } finally { db.close(); }
  await request('/tasks', {
    method: 'POST', user: 'local-viewer', status: 403,
    body: { title: 'Forbidden follow-up', topicId: 'launch', parentTaskId: parent.id, status: 'Not started' },
  });
});

test('parent validation rejects missing and cross-project parents and wrong topics', async () => {
  const parent = await createTask('Parent for validation');
  for (const extra of [{ parentTaskId: 'missing' }, { parentTaskId: {} }, { parentTaskId: parent.id, topicId: 'product' }]) {
    await request('/tasks', {
      method: 'POST', status: 400,
      body: { title: 'Invalid follow-up', topicId: 'launch', status: 'Not started', ...extra },
    });
  }
  const otherResponse = await fetch(`${baseUrl}/api/projects`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: 'Other project' }),
  });
  assert.equal(otherResponse.status, 201);
  const { project } = await otherResponse.json();
  const response = await fetch(`${baseUrl}/api/projects/${project.id}/tasks`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ title: 'Cross-project follow-up', topicId: project.topics[0].id, parentTaskId: parent.id, status: 'Not started' }),
  });
  assert.equal(response.status, 400);
});

test('nested follow-ups move together and retain links across deletion and restoration', async () => {
  const parent = await createTask('Movable parent');
  const child = await createTask('Child', { parentTaskId: parent.id });
  const grandchild = await createTask('Grandchild', { parentTaskId: child.id });
  await request(`/tasks/${parent.id}`, { method: 'PATCH', body: { parentTaskId: grandchild.id }, status: 400 });
  await request(`/tasks/${child.id}`, { method: 'PATCH', body: { topicId: 'product' }, status: 400 });
  await request(`/tasks/${parent.id}`, { method: 'PATCH', body: { topicId: 'product' } });
  let state = (await request('')).project;
  for (const id of [parent.id, child.id, grandchild.id]) assert.equal(state.tasks.find((task) => task.id === id).topicId, 'product');
  await request(`/tasks/${parent.id}`, { method: 'DELETE', status: 204 });
  state = (await request('')).project;
  assert.equal(state.tasks.find((task) => task.id === child.id).parentTaskId, parent.id);
  assert.ok(state.tasks.some((task) => task.id === grandchild.id));
  await request(`/recycle-bin/${parent.id}/restore`, { method: 'POST', body: { topicId: 'launch' }, status: 409 });
  await request(`/recycle-bin/${parent.id}/restore`, { method: 'POST', body: { topicId: 'product' } });
  await request(`/tasks/${child.id}`, { method: 'DELETE', status: 204 });
  const restored = await request(`/recycle-bin/${child.id}/restore`, { method: 'POST', body: { topicId: 'launch' } });
  assert.equal(restored.task.topicId, 'product');
  assert.equal(restored.task.parentTaskId, parent.id);
});
