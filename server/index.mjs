import { randomUUID } from 'node:crypto';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { fileURLToPath } from 'node:url';
import express from 'express';
import { createRemoteJWKSet, jwtVerify } from 'jose';

const root = dirname(fileURLToPath(import.meta.url));
const isProduction = process.env.NODE_ENV === 'production';
const dbPath = process.env.DB_PATH || join(root, '..', 'data', 'fieldnotes.db');
mkdirSync(dirname(dbPath), { recursive: true });
const db = new DatabaseSync(dbPath);
db.exec('PRAGMA journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS projects (id TEXT PRIMARY KEY, tenant_id TEXT NOT NULL, state TEXT NOT NULL, created_at TEXT NOT NULL);
  CREATE TABLE IF NOT EXISTS memberships (project_id TEXT NOT NULL, user_id TEXT NOT NULL, email TEXT NOT NULL, display_name TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('owner','manager','contributor','viewer')), PRIMARY KEY(project_id, user_id), FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
`);

const demoPeople = [
  { id: 'local-alex', email: 'alex@localhost', displayName: 'Alex Morgan', role: 'owner' },
  { id: 'local-maya', email: 'maya@localhost', displayName: 'Maya Chen', role: 'manager' },
  { id: 'local-jordan', email: 'jordan@localhost', displayName: 'Jordan Lee', role: 'contributor' },
  { id: 'local-viewer', email: 'viewer@localhost', displayName: 'Taylor Quinn', role: 'viewer' },
];

function seedDemo() {
  if (db.prepare('SELECT COUNT(*) AS count FROM projects WHERE tenant_id = ?').get('local-tenant').count) return;
  const projectId = 'northstar';
  const state = {
    id: projectId,
    name: 'Northstar launch',
    description: 'Bring the new customer workspace to market with a calm, confident launch.',
    topics: [
      { id: 'launch', name: 'Launch plan', parentId: null },
      { id: 'research', name: 'Customer research', parentId: 'launch' },
      { id: 'product', name: 'Product readiness', parentId: null },
      { id: 'comms', name: 'Launch communications', parentId: null },
    ],
    tasks: [
      { id: 't1', topicId: 'launch', title: 'Confirm launch scope with the team', description: 'Lock the first-release scope and record any deliberate exclusions.', dueDate: '2026-11-12', status: 'In progress', owner: 'Maya Chen', comments: ['Design review is booked for Thursday.'] },
      { id: 't2', topicId: 'launch', title: 'Publish the milestone calendar', description: 'Share one source of truth for launch dates and decision checkpoints.', dueDate: '2026-11-14', status: 'Not started', owner: 'Jordan Lee', comments: [] },
      { id: 't3', topicId: 'research', title: 'Schedule five customer interviews', description: 'Recruit a mix of existing and first-time customers.', dueDate: '2026-11-10', status: 'Blocked', owner: 'Maya Chen', comments: ['Waiting on the updated participant list.'] },
      { id: 't4', topicId: 'product', title: 'Close critical onboarding issues', description: 'Resolve the remaining P1 issues before release candidate review.', dueDate: '2026-11-18', status: 'In progress', owner: 'Sam Rivera', comments: [] },
      { id: 't5', projectId: 'northstar', topicId: 'comms', title: 'Review launch announcement draft', description: 'Check the customer-facing copy against the final feature list.', dueDate: '2026-11-20', status: 'Done', owner: 'Jordan Lee', comments: ['Reviewed and approved.'] },
    ],
  };
  const seed = () => {
    db.exec('BEGIN IMMEDIATE');
    try {
    db.prepare('INSERT INTO projects (id, tenant_id, state, created_at) VALUES (?, ?, ?, ?)').run(projectId, 'local-tenant', JSON.stringify(state), new Date().toISOString());
    const addMember = db.prepare('INSERT INTO memberships (project_id, user_id, email, display_name, role) VALUES (?, ?, ?, ?, ?)');
    for (const person of demoPeople) addMember.run(projectId, person.id, person.email, person.displayName, person.role);
      db.exec('COMMIT');
    } catch (error) {
      db.exec('ROLLBACK');
      throw error;
    }
  };
  seed();
}

if (!isProduction) seedDemo();

async function authenticate(req, res, next) {
  if (!isProduction) {
    const userId = String(req.header('x-dev-user') || 'local-alex');
    const person = demoPeople.find((item) => item.id === userId) ?? demoPeople[0];
    req.user = { id: person.id, email: person.email, name: person.displayName, tenantId: 'local-tenant', demo: true };
    return next();
  }

  const token = req.header('authorization')?.replace(/^Bearer\s+/i, '');
  const tenantId = process.env.ENTRA_TENANT_ID;
  const clientId = process.env.ENTRA_CLIENT_ID;
  if (!token || !tenantId || !clientId) return res.status(401).json({ error: 'Teams sign-in is not configured.' });
  try {
    const jwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`));
    const appIdUri = process.env.ENTRA_APP_ID_URI || `api://${process.env.TEAMS_DOMAIN}/${clientId}`;
    const { payload } = await jwtVerify(token, jwks, {
      issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
      audience: [clientId, appIdUri],
    });
    if (payload.tid !== tenantId || typeof payload.oid !== 'string') throw new Error('Invalid tenant or user claim');
    req.user = {
      id: payload.oid,
      email: String(payload.preferred_username || payload.email || '').toLowerCase(),
      name: String(payload.name || payload.preferred_username || 'Teams user'),
      tenantId,
      demo: false,
    };
    next();
  } catch {
    res.status(401).json({ error: 'Teams sign-in could not be verified.' });
  }
}

function getMembership(projectId, user) {
  return db.prepare('SELECT * FROM memberships WHERE project_id = ? AND (user_id = ? OR email = ?) ORDER BY CASE WHEN user_id = ? THEN 0 ELSE 1 END LIMIT 1')
    .get(projectId, user.id, user.email, user.id);
}

function requireProject(minimumRoles = ['owner', 'manager', 'contributor', 'viewer']) {
  return (req, res, next) => {
    const record = db.prepare('SELECT * FROM projects WHERE id = ? AND tenant_id = ?').get(req.params.projectId, req.user.tenantId);
    const membership = record && getMembership(record.id, req.user);
    if (!record || !membership) return res.status(404).json({ error: 'Project not found.' });
    if (!minimumRoles.includes(membership.role)) return res.status(403).json({ error: 'Your project role does not allow this action.' });
    req.projectRecord = record;
    req.projectState = JSON.parse(record.state);
    req.membership = membership;
    next();
  };
}

function saveProject(projectId, state) {
  db.prepare('UPDATE projects SET state = ? WHERE id = ?').run(JSON.stringify(state), projectId);
}

const app = express();
app.use(express.json({ limit: '1mb' }));
app.use('/api', authenticate);
app.get('/api/session', (req, res) => res.json(req.user));

app.get('/api/projects', (req, res) => {
  const rows = db.prepare(`SELECT p.id, p.state, m.role FROM projects p JOIN memberships m ON m.project_id = p.id
    WHERE p.tenant_id = ? AND (m.user_id = ? OR m.email = ?) ORDER BY p.created_at DESC`)
    .all(req.user.tenantId, req.user.id, req.user.email);
  res.json({ projects: rows.map((row) => {
    const state = JSON.parse(row.state);
    return { id: state.id, name: state.name, description: state.description, role: row.role };
  }) });
});

app.post('/api/projects', (req, res) => {
  const name = String(req.body.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Project name is required.' });
  const id = randomUUID();
  const topic = { id: randomUUID(), name: 'Getting started', parentId: null };
  const state = { id, name, description: String(req.body.description || ''), topics: [topic], tasks: [] };
  const create = () => {
    db.exec('BEGIN IMMEDIATE');
    try {
    db.prepare('INSERT INTO projects (id, tenant_id, state, created_at) VALUES (?, ?, ?, ?)').run(id, req.user.tenantId, JSON.stringify(state), new Date().toISOString());
    db.prepare('INSERT INTO memberships (project_id, user_id, email, display_name, role) VALUES (?, ?, ?, ?, ?)').run(id, req.user.id, req.user.email, req.user.name, 'owner');
      db.exec('COMMIT');
    } catch (error) {
      db.exec('ROLLBACK');
      throw error;
    }
  };
  create();
  res.status(201).json({ project: state, role: 'owner' });
});

app.get('/api/projects/:projectId', requireProject(), (req, res) => {
  res.json({ project: req.projectState, role: req.membership.role });
});

app.delete('/api/projects/:projectId', requireProject(['owner']), (req, res) => {
  db.exec('BEGIN IMMEDIATE');
  try {
    db.prepare('DELETE FROM memberships WHERE project_id = ?').run(req.params.projectId);
    db.prepare('DELETE FROM projects WHERE id = ? AND tenant_id = ?').run(req.params.projectId, req.user.tenantId);
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    throw error;
  }
  res.status(204).end();
});

app.patch('/api/projects/:projectId', requireProject(['owner', 'manager']), (req, res) => {
  const state = req.projectState;
  if (typeof req.body.name === 'string') {
    if (!req.body.name.trim()) return res.status(400).json({ error: 'Project name cannot be empty.' });
    state.name = req.body.name.trim();
  }
  if (typeof req.body.description === 'string') state.description = req.body.description;
  if (Array.isArray(req.body.topics)) state.topics = req.body.topics;
  saveProject(req.params.projectId, state);
  res.json({ project: state });
});

app.delete('/api/projects/:projectId/topics/:topicId', requireProject(['owner', 'manager']), (req, res) => {
  const topic = req.projectState.topics.find((item) => item.id === req.params.topicId);
  if (!topic) return res.status(404).json({ error: 'Topic not found.' });

  const removedTopicIds = new Set([topic.id]);
  let addedChild = true;
  while (addedChild) {
    addedChild = false;
    for (const item of req.projectState.topics) {
      if (item.parentId && removedTopicIds.has(item.parentId) && !removedTopicIds.has(item.id)) {
        removedTopicIds.add(item.id);
        addedChild = true;
      }
    }
  }
  if (removedTopicIds.size === req.projectState.topics.length) {
    return res.status(409).json({ error: 'A project must keep at least one topic.' });
  }

  const topicNames = new Map(req.projectState.topics.map((item) => [item.id, item.name]));
  const removedTasks = req.projectState.tasks.filter((task) => removedTopicIds.has(task.topicId));
  req.projectState.tasks = req.projectState.tasks.filter((task) => !removedTopicIds.has(task.topicId));
  req.projectState.recycleBin ??= [];
  req.projectState.recycleBin.unshift(...removedTasks.map((task) => ({
    ...task,
    topicName: topicNames.get(task.topicId) ?? topic.name,
    deletedAt: new Date().toISOString(),
  })));
  req.projectState.topics = req.projectState.topics.filter((item) => !removedTopicIds.has(item.id));
  saveProject(req.params.projectId, req.projectState);
  res.json({ removedTopicIds: [...removedTopicIds], archivedTaskCount: removedTasks.length });
});

app.post('/api/projects/:projectId/tasks', requireProject(['owner', 'manager', 'contributor']), (req, res) => {
  const { title, topicId, status } = req.body;
  if (!String(title || '').trim() || !req.projectState.topics.some((topic) => topic.id === topicId)) return res.status(400).json({ error: 'A title and valid topic are required.' });
  if (!['Not started', 'In progress', 'Blocked', 'Done'].includes(status)) return res.status(400).json({ error: 'Invalid task status.' });
  const task = { id: randomUUID(), topicId, title: String(title).trim(), description: String(req.body.description || ''), dueDate: String(req.body.dueDate || ''), status, owner: String(req.body.owner || ''), comments: [] };
  req.projectState.tasks.unshift(task);
  saveProject(req.params.projectId, req.projectState);
  res.status(201).json({ task });
});

app.patch('/api/projects/:projectId/tasks/:taskId', requireProject(['owner', 'manager', 'contributor']), (req, res) => {
  const task = req.projectState.tasks.find((item) => item.id === req.params.taskId);
  if (!task) return res.status(404).json({ error: 'Task not found.' });
  const allowedFields = req.membership.role === 'contributor' ? ['status', 'comments'] : ['title', 'description', 'dueDate', 'status', 'owner', 'topicId', 'comments'];
  for (const field of allowedFields) if (field in req.body) task[field] = req.body[field];
  if (task.title !== undefined && !String(task.title).trim()) return res.status(400).json({ error: 'Task title cannot be empty.' });
  if (task.status !== undefined && !['Not started', 'In progress', 'Blocked', 'Done'].includes(task.status)) return res.status(400).json({ error: 'Invalid task status.' });
  saveProject(req.params.projectId, req.projectState);
  res.json({ task });
});

app.delete('/api/projects/:projectId/tasks/:taskId', requireProject(['owner', 'manager']), (req, res) => {
  const task = req.projectState.tasks.find((item) => item.id === req.params.taskId);
  if (!task) return res.status(404).json({ error: 'Task not found.' });
  const topicName = req.projectState.topics.find((topic) => topic.id === task.topicId)?.name ?? 'Removed topic';
  req.projectState.tasks = req.projectState.tasks.filter((item) => item.id !== req.params.taskId);
  req.projectState.recycleBin ??= [];
  req.projectState.recycleBin.unshift({ ...task, topicName, deletedAt: new Date().toISOString() });
  saveProject(req.params.projectId, req.projectState);
  res.status(204).end();
});

app.post('/api/projects/:projectId/recycle-bin/:taskId/restore', requireProject(['owner', 'manager']), (req, res) => {
  const recycleBin = req.projectState.recycleBin ?? [];
  const index = recycleBin.findIndex((task) => task.id === req.params.taskId);
  if (index === -1) return res.status(404).json({ error: 'Recycle-bin task not found.' });
  const deletedTask = recycleBin[index];
  const requestedTopicId = String(req.body.topicId || deletedTask.topicId);
  if (!req.projectState.topics.some((topic) => topic.id === requestedTopicId)) {
    return res.status(400).json({ error: 'Choose an existing topic for this task.' });
  }
  const { topicName: _topicName, deletedAt: _deletedAt, ...task } = deletedTask;
  task.topicId = requestedTopicId;
  req.projectState.tasks.unshift(task);
  req.projectState.recycleBin = recycleBin.filter((item) => item.id !== req.params.taskId);
  saveProject(req.params.projectId, req.projectState);
  res.json({ task });
});

app.delete('/api/projects/:projectId/recycle-bin/:taskId', requireProject(['owner', 'manager']), (req, res) => {
  const recycleBin = req.projectState.recycleBin ?? [];
  if (!recycleBin.some((task) => task.id === req.params.taskId)) return res.status(404).json({ error: 'Recycle-bin task not found.' });
  req.projectState.recycleBin = recycleBin.filter((task) => task.id !== req.params.taskId);
  saveProject(req.params.projectId, req.projectState);
  res.status(204).end();
});

app.get('/api/projects/:projectId/members', requireProject(), (req, res) => {
  const members = db.prepare('SELECT user_id AS id, email, display_name AS name, role FROM memberships WHERE project_id = ? ORDER BY display_name').all(req.params.projectId);
  res.json({ members });
});

app.post('/api/projects/:projectId/members', requireProject(['owner', 'manager']), (req, res) => {
  const email = String(req.body.email || '').trim().toLowerCase();
  const role = String(req.body.role || 'viewer');
  const allowedRoles = req.membership.role === 'owner' ? ['owner', 'manager', 'contributor', 'viewer'] : ['contributor', 'viewer'];
  if (!email.includes('@') || !allowedRoles.includes(role)) {
    return res.status(400).json({ error: 'Enter a valid email and allowed role.' });
  }
  const id = email;
  db.prepare(`INSERT INTO memberships (project_id, user_id, email, display_name, role) VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(project_id, user_id) DO UPDATE SET role = excluded.role`).run(req.params.projectId, id, email, email.split('@')[0], role);
  res.status(201).json({ member: { id, email, name: email.split('@')[0], role } });
});

app.patch('/api/projects/:projectId/members/:userId', requireProject(['owner']), (req, res) => {
  const role = String(req.body.role || '');
  if (!['owner', 'manager', 'contributor', 'viewer'].includes(role)) return res.status(400).json({ error: 'Invalid member role.' });
  const current = db.prepare('SELECT role FROM memberships WHERE project_id = ? AND user_id = ?').get(req.params.projectId, req.params.userId);
  if (!current) return res.status(404).json({ error: 'Member not found.' });
  if (current.role === 'owner' && role !== 'owner' && db.prepare('SELECT COUNT(*) AS count FROM memberships WHERE project_id = ? AND role = \'owner\'').get(req.params.projectId).count < 2) {
    return res.status(409).json({ error: 'A project must keep at least one owner.' });
  }
  const result = db.prepare('UPDATE memberships SET role = ? WHERE project_id = ? AND user_id = ?').run(role, req.params.projectId, req.params.userId);
  res.json({ ok: true });
});

app.delete('/api/projects/:projectId/members/:userId', requireProject(['owner']), (req, res) => {
  const member = db.prepare('SELECT role FROM memberships WHERE project_id = ? AND user_id = ?').get(req.params.projectId, req.params.userId);
  if (!member) return res.status(404).json({ error: 'Member not found.' });
  if (member.role === 'owner' && db.prepare('SELECT COUNT(*) AS count FROM memberships WHERE project_id = ? AND role = \'owner\'').get(req.params.projectId).count < 2) {
    return res.status(409).json({ error: 'A project must keep at least one owner.' });
  }
  db.prepare('DELETE FROM memberships WHERE project_id = ? AND user_id = ?').run(req.params.projectId, req.params.userId);
  res.status(204).end();
});

if (isProduction) {
  const dist = join(root, '..', 'dist');
  app.use(express.static(dist));
  app.get('*', (_req, res) => res.sendFile(join(dist, 'index.html')));
}

const port = Number(process.env.PORT || 3001);
app.listen(port, '0.0.0.0', () => console.log(`Fieldnotes API listening on ${port}`));