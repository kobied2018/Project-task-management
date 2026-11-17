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

test('legacy tasks and recycled tasks receive persistent numbers without changing links or content', async () => {
  const db = new DatabaseSync(join(directory, 'test.db'));
  let original;
  try {
    original = JSON.parse(db.prepare('SELECT state FROM projects WHERE id = ?').get('northstar').state);
    original.tasks[0].comments.push('[Milestone calendar](task:t2)');
    original.recycleBin = [{ ...original.tasks[0], id: 'legacy-removed', parentTaskId: 't1', topicName: 'Launch plan', deletedAt: '2026-11-01T00:00:00Z' }];
    db.prepare('UPDATE projects SET state = ? WHERE id = ?').run(JSON.stringify(original), 'northstar');
    const { project } = await request('');
    const withoutNumber = ({ taskNumber, ...task }) => task;
    assert.deepEqual(project.tasks.map(withoutNumber), original.tasks);
    assert.deepEqual(project.recycleBin.map(withoutNumber), original.recycleBin);
    assert.deepEqual([...project.tasks, ...project.recycleBin].map((task) => task.taskNumber), [1, 2, 3, 4, 5, 6]);
    assert.equal(project.nextTaskNumber, 7);
    assert.deepEqual((await request('')).project, project);
    const saved = JSON.parse(db.prepare('SELECT state FROM projects WHERE id = ?').get('northstar').state);
    assert.deepEqual(saved, project);
  } finally { db.close(); }
});

test('follow-ups persist with independent fields, comments, and permissions', async () => {
  const parent = await createTask('Completed parent', { status: 'Done' });
  const child = await createTask('Follow up after completion', {
    parentTaskId: parent.id, owner: 'Test owner', dueDate: '2026-12-01',
  }, 'local-jordan');
  assert.equal(child.taskNumber, parent.taskNumber + 1);
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

test('task numbering survives edits, deletion and restoration, and simultaneous creates', async () => {
  const parent = await createTask('Numbered parent', { taskNumber: 9000 });
  const child = await createTask('Numbered child', { parentTaskId: parent.id });
  assert.notEqual(parent.taskNumber, 9000);
  const updated = (await request(`/tasks/${child.id}`, {
    method: 'PATCH', body: { title: 'Renamed child', taskNumber: 1, id: 'changed-id', status: 'Done' },
  })).task;
  assert.equal(updated.taskNumber, child.taskNumber);
  assert.equal(updated.id, child.id);
  await request(`/tasks/${child.id}`, { method: 'DELETE', status: 204 });
  const whileRemoved = await createTask('Created while child is recycled');
  assert.equal(whileRemoved.taskNumber, child.taskNumber + 1);
  const restored = (await request(`/recycle-bin/${child.id}/restore`, { method: 'POST', body: {} })).task;
  assert.equal(restored.taskNumber, child.taskNumber);
  assert.equal(restored.parentTaskId, parent.id);
  const reopened = (await request(`/tasks/${child.id}`, { method: 'PATCH', body: { status: 'In progress' } })).task;
  assert.equal(reopened.taskNumber, child.taskNumber);
  await request(`/tasks/${whileRemoved.id}`, { method: 'DELETE', status: 204 });
  await request(`/recycle-bin/${whileRemoved.id}`, { method: 'DELETE', status: 204 });
  const concurrent = await Promise.all(Array.from({ length: 5 }, (_, index) => createTask(`Concurrent numbered task ${index}`)));
  assert.deepEqual(concurrent.map((task) => task.taskNumber).sort((a, b) => a - b),
    Array.from({ length: 5 }, (_, index) => whileRemoved.taskNumber + index + 1));
  const persisted = (await request('')).project.tasks;
  for (const task of concurrent) assert.equal(persisted.find((item) => item.id === task.id).taskNumber, task.taskNumber);

  const response = await fetch(`${baseUrl}/api/projects`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: 'Independent numbering' }),
  });
  assert.equal(response.status, 201);
  const { project } = await response.json();
  const created = await fetch(`${baseUrl}/api/projects/${project.id}/tasks`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ title: 'First task in another project', topicId: project.topics[0].id, status: 'Not started' }),
  });
  assert.equal(created.status, 201);
  assert.equal((await created.json()).task.taskNumber, 1);
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

test('detailed descriptions persist independently of summaries and respect editing permissions', async () => {
  const task = await createTask('Task with collapsible details', {
    description: 'Always-visible summary',
    detailedDescription: 'Problem:\nOriginal issue\n\nImpact:\nPotential effect\n\nRecommendation:\nProposed action',
  });
  const details = `${task.detailedDescription}\n\nClosure evidence:\nReviewed output`;
  const updated = (await request(`/tasks/${task.id}`, { method: 'PATCH', body: { detailedDescription: details } })).task;
  assert.deepEqual(updated, { ...task, detailedDescription: details });
  const contributor = (await request(`/tasks/${task.id}`, {
    method: 'PATCH', user: 'local-jordan',
    body: { description: 'Unauthorized summary', detailedDescription: 'Unauthorized details', status: 'In progress' },
  })).task;
  assert.deepEqual(contributor, { ...updated, status: 'In progress' });
  await request(`/tasks/${task.id}`, {
    method: 'PATCH', user: 'local-viewer', status: 403, body: { detailedDescription: 'Forbidden' },
  });
  await request(`/tasks/${task.id}`, { method: 'PATCH', status: 400, body: { detailedDescription: [] } });
  await request('/tasks', {
    method: 'POST', status: 400,
    body: { title: 'Invalid details', topicId: 'launch', status: 'Not started', detailedDescription: {} },
  });
  const stored = (await request('')).project.tasks.find((item) => item.id === task.id);
  assert.deepEqual(stored, contributor);
  await request(`/tasks/${task.id}`, { method: 'DELETE', status: 204 });
  const restored = (await request(`/recycle-bin/${task.id}/restore`, { method: 'POST', body: {} })).task;
  assert.deepEqual(restored, stored);
  const cleared = (await request(`/tasks/${task.id}`, { method: 'PATCH', body: { detailedDescription: '' } })).task;
  assert.equal(cleared.detailedDescription, '');
  assert.equal(cleared.description, task.description);
});
