import { useEffect, useMemo, useState } from 'react';
import { app as teamsApp, authentication } from '@microsoft/teams-js';
import {
  AlignLeft, Archive, ArrowDown, ArrowUp, CalendarDays, Check, ChevronDown, ChevronRight, Copy,
  CircleHelp, ExternalLink, FileUp, FolderKanban, ListFilter, MessageSquareText, MoreHorizontal,
  Pencil, Plus, RotateCcw, Search, ShieldCheck, Sparkles, Trash2, Type, UsersRound, X,
} from 'lucide-react';

type Role = 'owner' | 'manager' | 'contributor' | 'viewer';
type Status = 'Not started' | 'In progress' | 'Blocked' | 'Done';
type Topic = { id: string; name: string; parentId: string | null };
type Task = { id: string; topicId: string; title: string; description: string; dueDate: string; status: Status; owner: string; comments: string[] };
type DeletedTask = Task & { topicName: string; deletedAt: string };
type Project = { id: string; name: string; description: string; topics: Topic[]; tasks: Task[]; recycleBin?: DeletedTask[] };
type ProjectSummary = { id: string; name: string; description: string; role: Role };
type Member = { id: string; email: string; name: string; role: Role };
type Candidate = Omit<Task, 'id' | 'topicId' | 'comments'> & { selected: boolean };
type CopilotProposal =
  | { action: 'create'; topicId: string; title: string; description: string; dueDate: string; status: Status; owner: string; selected: boolean }
  | { action: 'update'; taskId: string; taskTitle: string; changes: Partial<Pick<Task, 'title' | 'description' | 'dueDate' | 'status' | 'owner'>>; selected: boolean };
type RemovalTarget =
  | { kind: 'project'; project: ProjectSummary }
  | { kind: 'topic'; topic: Topic }
  | { kind: 'task'; task: Task }
  | { kind: 'deleted-task'; task: DeletedTask };

const statuses: Status[] = ['Not started', 'In progress', 'Blocked', 'Done'];
const demoUsers = [
  { id: 'local-alex', name: 'Alex Morgan', role: 'Owner' },
  { id: 'local-maya', name: 'Maya Chen', role: 'Manager' },
  { id: 'local-jordan', name: 'Jordan Lee', role: 'Contributor' },
  { id: 'local-viewer', name: 'Taylor Quinn', role: 'Viewer' },
];

async function request<T>(path: string, userId: string, token: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers);
  if (init?.body) headers.set('Content-Type', 'application/json');
  if (import.meta.env.DEV) headers.set('x-dev-user', userId);
  if (token) headers.set('Authorization', `Bearer ${token}`);
  const response = await fetch(path, { ...init, headers });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(body.error || `Request failed (${response.status})`);
  }
  return response.status === 204 ? undefined as T : response.json();
}

function parseTranscript(text: string): Candidate[] {
  const candidateLine = /\b(todo|action item|follow[- ]?up|will|needs? to|should|must)\b/i;
  const prefix = /^\s*(?:[-*]\s*)?(?:\[?\s*(?:todo|action item)\s*\]?\s*[:\-]?\s*)/i;
  return text.split(/\r?\n/).map((line) => line.trim()).filter((line) => candidateLine.test(line) && line.length > 8).map((line) => {
    const owner = line.match(/\b(?:owner|assigned to)\s*:?\s*([A-Z][\w.-]+(?:\s+[A-Z][\w.-]+){0,2})/i)?.[1] ?? '';
    const dueDate = line.match(/\b(20\d{2}-\d{2}-\d{2})\b/)?.[1] ?? '';
    const explicitStatus = line.match(/\bstatus\s*:?[\t ]*(not started|in progress|blocked|completed|done)\b/i)?.[1]?.toLowerCase();
    const match = explicitStatus ?? line.match(/\b(blocked|in progress|completed|done)\b/i)?.[1]?.toLowerCase();
    const status: Status = match === 'done' || match === 'completed' ? 'Done' : match === 'blocked' ? 'Blocked' : match === 'in progress' ? 'In progress' : 'Not started';
    let title = line.replace(prefix, '').replace(/\b(?:owner|assigned to)\s*:?.*$/i, '').replace(/\s*\(?20\d{2}-\d{2}-\d{2}\)?\s*/i, ' ').trim();
    title = title.replace(/^(?:we\s+)?(?:will|need to|needs to|should|must)\s+/i, '');
    title = title.replace(/\b(?:by|on|due)\s*[,.;]?\s*$/i, '').replace(/[.;,\s]+$/, '');
    return { title: title.charAt(0).toUpperCase() + title.slice(1), description: `From transcript: “${line}”`, dueDate, status, owner, selected: true };
  }).filter((candidate) => candidate.title.length > 4);
}

function parseCopilotProposals(text: string, project: Project): CopilotProposal[] | null {
  const jsonText = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  try {
    const parsed = JSON.parse(jsonText) as { proposals?: unknown };
    if (!Array.isArray(parsed.proposals)) return null;
    const proposals: CopilotProposal[] = [];
    for (const raw of parsed.proposals) {
      if (!raw || typeof raw !== 'object') continue;
      const item = raw as Record<string, unknown>;
      if (item.action === 'create') {
        const title = typeof item.title === 'string' ? item.title.trim() : '';
        const topicId = typeof item.topicId === 'string' ? item.topicId : '';
        if (!title || !project.topics.some((topic) => topic.id === topicId)) continue;
        proposals.push({
          action: 'create',
          topicId,
          title,
          description: typeof item.description === 'string' ? item.description : '',
          dueDate: typeof item.dueDate === 'string' ? item.dueDate : '',
          status: statuses.includes(item.status as Status) ? item.status as Status : 'Not started',
          owner: typeof item.owner === 'string' ? item.owner : 'Unassigned',
          selected: true,
        });
      } else if (item.action === 'update' && typeof item.taskId === 'string' && item.changes && typeof item.changes === 'object') {
        const task = project.tasks.find((candidate) => candidate.id === item.taskId);
        if (!task || task.status === 'Done') continue;
        const rawChanges = item.changes as Record<string, unknown>;
        const validated: Partial<Pick<Task, 'title' | 'description' | 'dueDate' | 'status' | 'owner'>> = {};
        for (const field of ['title', 'description', 'dueDate', 'owner'] as const) {
          if (typeof rawChanges[field] === 'string') validated[field] = rawChanges[field] as string;
        }
        if (statuses.includes(rawChanges.status as Status)) validated.status = rawChanges.status as Status;
        if (!Object.keys(validated).length) continue;
        proposals.push({ action: 'update', taskId: task.id, taskTitle: task.title, changes: validated, selected: true });
      }
    }
    return proposals;
  } catch {
    return null;
  }
}

function buildCopilotTranscriptPrompt(transcript: string, project: Project, targetTopicId: string) {
  const topicList = project.topics.map((topic) => ({ id: topic.id, name: topic.name, parentId: topic.parentId }));
  const taskList = project.tasks.map((task) => ({
    id: task.id,
    topicId: task.topicId,
    title: task.title,
    description: task.description,
    dueDate: task.dueDate,
    status: task.status,
    owner: task.owner,
    comments: task.comments,
  }));
  return `Review this Teams meeting transcript against the current Fieldnotes project. Propose only changes clearly supported by the meeting. Do not make changes directly; return proposals for a user to review.

Project: ${project.name} (${project.id})
Default topic for new tasks: ${targetTopicId}

Current topics:
${JSON.stringify(topicList, null, 2)}

Current tasks (use these exact task IDs for updates):
${JSON.stringify(taskList, null, 2)}

Return valid JSON only, in this schema:
{
  "proposals": [
    {
      "action": "update",
      "taskId": "existing-task-id",
      "changes": { "status": "Done" }
    },
    {
      "action": "create",
      "topicId": "existing-topic-id",
      "title": "Concise action title",
      "description": "Brief context or acceptance criteria",
      "dueDate": "YYYY-MM-DD or empty string",
      "status": "Not started",
      "owner": "Full Name or Unassigned"
    }
  ]
}

Rules:
- Compare the transcript with existing tasks. Propose an update only when the meeting clearly reports new progress, a changed owner/date/description, or an outcome.
- Use status Done only when completion is explicit; use Blocked or In progress only when clearly stated. Do not change status based on speculation.
- For updates, include only fields that should change. Use the exact existing task ID; never invent IDs. Do not propose updates to tasks already marked Done.
- Create a new task only for a distinct action not already represented by an existing task. Use the default topic ID unless the transcript clearly maps it to another listed topic ID.
- Include a due date only when stated with an unambiguous year. Otherwise use an empty string for new tasks and omit dueDate from updates.
- Use Unassigned when no owner is clearly identified. Never infer a person from who is speaking.
- Preserve existing task details and fields that do not need to change.
- Merge duplicate actions. Do not include transcript dialogue, summaries, markdown fences, or commentary.
- If nothing should be created or updated, return exactly: {"proposals": []}

TRANSCRIPT:
${transcript}`;
}

function stripVttMetadata(text: string) {
  return text
    .replace(/^\uFEFF?WEBVTT[^\r\n]*\r?\n/i, '')
    .replace(/^\s*\d+\s*$/gm, '')
    .replace(/^\s*\d{2}:\d{2}(?::\d{2})?[.,]\d{3}\s+-->.*$/gm, '')
    .replace(/^\s*(?:NOTE|STYLE|REGION)\b.*$/gm, '')
    .replace(/<[^>]*>/g, '')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\n{2,}/g, '\n')
    .trim();
}

function roleTitle(role: Role) {
  return role.charAt(0).toUpperCase() + role.slice(1);
}

function dateLabel(value: string) {
  if (!value) return 'No date';
  return new Intl.DateTimeFormat('en', { month: 'short', day: 'numeric' }).format(new Date(`${value}T12:00:00`));
}

function statusKey(status: Status) {
  return status.toLowerCase().replaceAll(' ', '-');
}

export default function App() {
  const [devUser, setDevUser] = useState('local-alex');
  const [token, setToken] = useState('');
  const [session, setSession] = useState<{ id: string; name: string; email: string; demo: boolean } | null>(null);
  const [projects, setProjects] = useState<ProjectSummary[]>([]);
  const [project, setProject] = useState<Project | null>(null);
  const [role, setRole] = useState<Role>('viewer');
  const [activeTopicId, setActiveTopicId] = useState<string | null>(null);
  const [expanded, setExpanded] = useState<string[]>(['launch']);
  const [filter, setFilter] = useState<'All tasks' | Status>('All tasks');
  const [search, setSearch] = useState('');
  const [taskFontSize, setTaskFontSize] = useState(16);
  const [commentDrafts, setCommentDrafts] = useState<Record<string, string>>({});
  const [dialog, setDialog] = useState<'project' | 'edit-project' | 'topic' | 'edit-topic' | 'task' | 'transcript' | 'members' | null>(null);
  const [editingTask, setEditingTask] = useState<Task | null>(null);
  const [editingTopic, setEditingTopic] = useState<Topic | null>(null);
  const [pendingRemoval, setPendingRemoval] = useState<RemovalTarget | null>(null);
  const [showRecycleBin, setShowRecycleBin] = useState(false);
  const [notice, setNotice] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(true);
  const [candidates, setCandidates] = useState<Candidate[]>([]);
  const [copilotProposals, setCopilotProposals] = useState<CopilotProposal[] | null>(null);
  const [transcriptName, setTranscriptName] = useState('');
  const [transcriptText, setTranscriptText] = useState('');
  const [copilotPrompt, setCopilotPrompt] = useState('');
  const [busy, setBusy] = useState(false);
  const [projectMenu, setProjectMenu] = useState(false);

  async function loadProject(projectId: string, userId = devUser, accessToken = token) {
    const result = await request<{ project: Project; role: Role }>(`/api/projects/${projectId}`, userId, accessToken);
    setProject(result.project);
    setRole(result.role);
  }

  async function loadWorkspace(userId = devUser, accessToken = token) {
    setLoading(true);
    setError('');
    try {
      let nextToken = accessToken;
      if (!import.meta.env.DEV && !nextToken) {
        await teamsApp.initialize();
        nextToken = await authentication.getAuthToken();
        setToken(nextToken);
      }
      const currentSession = await request<{ id: string; name: string; email: string; demo: boolean }>('/api/session', userId, nextToken);
      const list = await request<{ projects: ProjectSummary[] }>('/api/projects', userId, nextToken);
      setSession(currentSession);
      setProjects(list.projects);
      if (list.projects.length) {
        const selected = list.projects.find((item) => item.id === project?.id) ?? list.projects[0];
        await loadProject(selected.id, userId, nextToken);
      } else {
        setProject(null);
      }
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Could not connect to the project workspace.');
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { void loadWorkspace(devUser); }, [devUser]);

  const canManage = role === 'owner' || role === 'manager';
  const canManageMembers = canManage;
  const canEditWork = role !== 'viewer';
  const currentTopic = project?.topics.find((topic) => topic.id === activeTopicId);
  const projectTasks = project?.tasks ?? [];
  const topicTasks = projectTasks.filter((task) => !activeTopicId || task.topicId === activeTopicId);
  const visibleTasks = useMemo(() => topicTasks.filter((task) => {
    const statusMatch = filter === 'All tasks' || task.status === filter;
    const searchMatch = !search || `${task.title} ${task.description} ${task.owner}`.toLowerCase().includes(search.toLowerCase());
    return statusMatch && searchMatch;
  }), [filter, search, project, activeTopicId]);
  const recycleBin = project?.recycleBin ?? [];
  const topicRecycleBin = recycleBin.filter((task) => !activeTopicId || task.topicId === activeTopicId);
  const visibleRecycleBin = useMemo(() => topicRecycleBin.filter((task) =>
    !search || `${task.title} ${task.topicName} ${task.owner}`.toLowerCase().includes(search.toLowerCase()),
  ), [project, activeTopicId, search]);
  const completed = projectTasks.filter((task) => task.status === 'Done').length;
  const blocked = projectTasks.filter((task) => task.status === 'Blocked').length;
  const completion = projectTasks.length ? Math.round(completed / projectTasks.length * 100) : 0;

  async function createProject(name: string, description: string) {
    const result = await request<{ project: Project; role: Role }>('/api/projects', devUser, token, { method: 'POST', body: JSON.stringify({ name, description }) });
    const list = await request<{ projects: ProjectSummary[] }>('/api/projects', devUser, token);
    setProjects(list.projects);
    setProject(result.project);
    setRole(result.role);
    setDialog(null);
    setActiveTopicId(result.project.topics[0]?.id ?? null);
    setNotice('Project created');
  }

  async function saveTopics(nextTopics: Topic[]) {
    if (!project) return;
    const result = await request<{ project: Project }>(`/api/projects/${project.id}`, devUser, token, { method: 'PATCH', body: JSON.stringify({ topics: nextTopics }) });
    setProject(result.project);
  }

  async function updateProject(name: string, description: string) {
    if (!project) return;
    const result = await request<{ project: Project }>(`/api/projects/${project.id}`, devUser, token, { method: 'PATCH', body: JSON.stringify({ name, description }) });
    setProject(result.project);
    setProjects((items) => items.map((item) => item.id === project.id ? { ...item, name, description } : item));
    setDialog(null);
    setNotice('Project updated');
  }

  async function updateTopic(topic: Topic, name: string, parentId: string | null) {
    if (!project) return;
    await saveTopics(project.topics.map((item) => item.id === topic.id ? { ...item, name, parentId } : item));
    setDialog(null);
    setNotice('Topic updated');
  }

  async function addTopic(name: string, parentId: string | null) {
    if (!project) return;
    const topic = { id: crypto.randomUUID(), name, parentId };
    await saveTopics([...project.topics, topic]);
    setActiveTopicId(topic.id);
    setExpanded((items) => parentId ? [...new Set([...items, parentId])] : items);
    setDialog(null);
  }

  async function moveTopic(topicId: string, direction: -1 | 1) {
    if (!project) return;
    const topic = project.topics.find((item) => item.id === topicId);
    if (!topic) return;
    const siblings = project.topics.filter((item) => item.parentId === topic.parentId);
    const index = siblings.findIndex((item) => item.id === topicId);
    const target = siblings[index + direction];
    if (!target) return;
    const next = [...project.topics];
    const from = next.findIndex((item) => item.id === topicId);
    const to = next.findIndex((item) => item.id === target.id);
    [next[from], next[to]] = [next[to], next[from]];
    try { await saveTopics(next); } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not reorder topics.'); }
  }

  async function addTask(task: Omit<Task, 'id'>) {
    if (!project) return;
    await request(`/api/projects/${project.id}/tasks`, devUser, token, { method: 'POST', body: JSON.stringify(task) });
    await loadProject(project.id);
    setDialog(null);
    setNotice('Task added');
  }

  async function updateTask(task: Task, changes: Partial<Task>) {
    if (!project) return;
    await request(`/api/projects/${project.id}/tasks/${task.id}`, devUser, token, { method: 'PATCH', body: JSON.stringify(changes) });
    await loadProject(project.id);
    setEditingTask(null);
    setDialog(null);
    setNotice('Task updated');
  }

  async function updateStatus(task: Task, status: Status) {
    if (!project) return;
    try {
      await request(`/api/projects/${project.id}/tasks/${task.id}`, devUser, token, { method: 'PATCH', body: JSON.stringify({ status }) });
      await loadProject(project.id);
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not update task.'); }
  }

  async function saveTaskComment(task: Task) {
    if (!project || !canEditWork || task.status === 'Done') return;
    const newComment = commentDrafts[task.id]?.trim();
    if (!newComment) return;
    const comments = [...task.comments, newComment];
    try {
      await request(`/api/projects/${project.id}/tasks/${task.id}`, devUser, token, { method: 'PATCH', body: JSON.stringify({ comments }) });
      await loadProject(project.id);
      setCommentDrafts((drafts) => {
        const next = { ...drafts };
        delete next[task.id];
        return next;
      });
      setNotice('Comment added');
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not save comments.'); }
  }

  async function restoreTask(task: DeletedTask) {
    if (!project) return;
    const topicId = project.topics.some((topic) => topic.id === task.topicId)
      ? task.topicId
      : activeTopicId ?? project.topics[0]?.id;
    if (!topicId) {
      setError('Add a topic before restoring this task.');
      return;
    }
    try {
      await request(`/api/projects/${project.id}/recycle-bin/${task.id}/restore`, devUser, token, { method: 'POST', body: JSON.stringify({ topicId }) });
      await loadProject(project.id);
      setNotice('Task restored');
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not restore task.'); }
  }

  async function confirmRemoval() {
    if (!pendingRemoval || !project) return;
    const removal = pendingRemoval;
    try {
      if (removal.kind === 'project') {
        await request(`/api/projects/${removal.project.id}`, devUser, token, { method: 'DELETE' });
        const remaining = projects.filter((item) => item.id !== removal.project.id);
        setProjects(remaining);
        setProjectMenu(false);
        if (project.id === removal.project.id) {
          setActiveTopicId(null);
          setShowRecycleBin(false);
          if (remaining.length) await loadProject(remaining[0].id);
          else setProject(null);
        }
        setNotice('Project removed');
      } else if (removal.kind === 'topic') {
        const descendants = new Set([removal.topic.id]);
        let changed = true;
        while (changed) {
          changed = false;
          for (const topic of project.topics) {
            if (topic.parentId && descendants.has(topic.parentId) && !descendants.has(topic.id)) {
              descendants.add(topic.id);
              changed = true;
            }
          }
        }
        const result = await request<{ archivedTaskCount: number }>(`/api/projects/${project.id}/topics/${removal.topic.id}`, devUser, token, { method: 'DELETE' });
        await loadProject(project.id);
        if (activeTopicId && descendants.has(activeTopicId)) setActiveTopicId(null);
        setExpanded((items) => items.filter((id) => !descendants.has(id)));
        setShowRecycleBin(true);
        setNotice(`${removal.topic.name} removed; ${result.archivedTaskCount} task${result.archivedTaskCount === 1 ? '' : 's'} moved to the recycle bin`);
      } else if (removal.kind === 'task') {
        await request(`/api/projects/${project.id}/tasks/${removal.task.id}`, devUser, token, { method: 'DELETE' });
        await loadProject(project.id);
        setShowRecycleBin(true);
        setNotice('Task moved to the topic recycle bin');
      } else {
        await request(`/api/projects/${project.id}/recycle-bin/${removal.task.id}`, devUser, token, { method: 'DELETE' });
        await loadProject(project.id);
        setNotice('Task permanently deleted');
      }
      setPendingRemoval(null);
      setDialog(null);
      setEditingTask(null);
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not remove item.'); }
  }

  async function importCandidates() {
    if (!project) return;
    setBusy(true);
    try {
      if (copilotProposals !== null) {
        const selectedProposals = copilotProposals.filter((proposal) => proposal.selected);
        if (role === 'contributor' && selectedProposals.some((proposal) =>
          proposal.action === 'update' && Object.keys(proposal.changes).some((field) => field !== 'status'),
        )) {
          throw new Error('Your role can update task status, but only Owners and Managers can change other task fields.');
        }
        for (const proposal of selectedProposals) {
          if (proposal.action === 'create') {
            await request(`/api/projects/${project.id}/tasks`, devUser, token, {
              method: 'POST',
              body: JSON.stringify({ topicId: proposal.topicId, title: proposal.title, description: proposal.description, dueDate: proposal.dueDate, status: proposal.status, owner: proposal.owner }),
            });
          } else {
            await request(`/api/projects/${project.id}/tasks/${proposal.taskId}`, devUser, token, { method: 'PATCH', body: JSON.stringify(proposal.changes) });
          }
        }
        await loadProject(project.id);
        setDialog(null);
        setCopilotProposals(null);
        setCandidates([]);
        setNotice(`${selectedProposals.length} proposed change${selectedProposals.length === 1 ? '' : 's'} applied`);
        return;
      }
      const topicId = activeTopicId ?? project.topics[0]?.id;
      if (!topicId) throw new Error('Add a topic before importing tasks.');
      const selected = candidates.filter((candidate) => candidate.selected);
      for (const candidate of selected) {
        const { selected: _selected, ...task } = candidate;
        await request(`/api/projects/${project.id}/tasks`, devUser, token, { method: 'POST', body: JSON.stringify({ ...task, topicId }) });
      }
      await loadProject(project.id);
      setDialog(null);
      setCandidates([]);
      setNotice(`${selected.length} task${selected.length === 1 ? '' : 's'} added from transcript`);
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not import tasks.'); }
    finally { setBusy(false); }
  }

  async function prepareWithCopilot() {
    if (!transcriptText.trim() || !project) {
      setError('Load a transcript before preparing it with Copilot.');
      return;
    }
    const targetTopicId = activeTopicId ?? project.topics[0]?.id;
    if (!targetTopicId) {
      setError('Add a topic before preparing tasks with Copilot.');
      return;
    }
    const prompt = buildCopilotTranscriptPrompt(transcriptText, project, targetTopicId);
    setCopilotPrompt(prompt);
    try {
      await navigator.clipboard.writeText(prompt);
      setNotice('Prompt copied. Open Microsoft 365 Copilot and paste it into chat.');
    } catch {
      setNotice('Prompt is ready below. Select and copy it, then paste it into Microsoft 365 Copilot.');
    }
  }

  async function copyCopilotPrompt() {
    try {
      await navigator.clipboard.writeText(copilotPrompt);
      setNotice('Copilot prompt copied');
    } catch {
      setNotice('Select the prompt text and copy it manually');
    }
  }

  function pasteCopilotResult() {
    setCopilotPrompt('');
    setTranscriptName('Copilot response');
    setTranscriptText('');
    setCandidates([]);
    setCopilotProposals(null);
  }

  function analyzeTranscriptResult(value: string) {
    if (!project) return;
    const proposals = parseCopilotProposals(value, project);
    setTranscriptText(value);
    setTranscriptName('Copilot response');
    setCopilotProposals(proposals);
    setCandidates(proposals === null ? parseTranscript(value) : []);
  }

  async function switchProject(projectId: string) {
    setProjectMenu(false);
    setActiveTopicId(null);
    setFilter('All tasks');
    setSearch('');
    try { await loadProject(projectId); } catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not load project.'); }
  }

  if (loading) return <main className="loading-screen"><div className="brand-mark"><span /><span /><span /></div><p>Opening your workspace…</p></main>;
  if (error && !session) return <main className="connection-screen"><div className="connection-mark"><ShieldCheck size={23} /></div><span className="eyebrow">FIELDNOTES · TEAMS</span><h1>Sign-in needed</h1><p>{error}</p><button className="button primary" onClick={() => void loadWorkspace(devUser, '')}>Try again</button></main>;

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand-row"><div className="brand-mark"><span /><span /><span /></div><span className="brand-name">fieldnotes</span></div>
      <div className="workspace-badge"><div className="workspace-icon">A</div><div><strong>Atlas Studio</strong><span>Teams workspace</span></div><MoreHorizontal size={17} /></div>
      <div className="sidebar-label-row"><span>PROJECTS</span><button aria-label="Create project" onClick={() => setDialog('project')}><Plus size={16} /></button></div>
      <div className="project-selector-wrap">
        {project && <button className="project-selector" onClick={() => setProjectMenu(!projectMenu)}><span className="project-dot" />{project.name}<ChevronDown size={14} /></button>}
        {projectMenu && <div className="project-menu">{projects.map((item) => <div className="project-menu-row" key={item.id}><button className="project-menu-select" onClick={() => void switchProject(item.id)}><span className="project-dot" />{item.name}<span>{roleTitle(item.role)}</span></button>{(item.role === 'owner' || item.role === 'manager') && <button className="icon-button" aria-label={`Edit ${item.name}`} title="Edit project" onClick={() => { void switchProject(item.id).then(() => setDialog('edit-project')); setProjectMenu(false); }}><Pencil size={14} /></button>}{item.role === 'owner' && <button className="icon-button danger-icon" aria-label={`Remove ${item.name}`} title="Remove project" onClick={() => { setProjectMenu(false); setPendingRemoval({ kind: 'project', project: item }); }}><Trash2 size={14} /></button>}</div>)}</div>}
      </div>
      <div className="sidebar-label-row topic-label"><span>PROJECT SPACE</span><button disabled={!canManage} aria-label="Add topic" onClick={() => { setEditingTopic(null); setDialog('topic'); }}><Plus size={16} /></button></div>
      <nav className="topic-tree" aria-label="Project topics">
        <button className={`tree-overview ${activeTopicId === null ? 'active' : ''}`} onClick={() => { setActiveTopicId(null); setFilter('All tasks'); }}><FolderKanban size={15} /><span>Project overview</span><b>{projectTasks.length}</b></button>
        {project?.topics.filter((topic) => !topic.parentId).map((topic) => <TopicBranch key={topic.id} topic={topic} topics={project.topics} tasks={projectTasks} activeId={activeTopicId} expanded={expanded} canManage={canManage} onSelect={(id) => { setActiveTopicId(id); setFilter('All tasks'); setShowRecycleBin(false); }} onExpand={(id) => setExpanded((items) => items.includes(id) ? items.filter((value) => value !== id) : [...items, id])} onAddChild={(id) => { setActiveTopicId(id); setEditingTopic(null); setDialog('topic'); }} onEdit={(item) => { setEditingTopic(item); setDialog('edit-topic'); }} onRemove={(item) => setPendingRemoval({ kind: 'topic', topic: item })} onMove={moveTopic} />)}
      </nav>
      <div className="sidebar-bottom">
        <div className="local-save"><span /> Shared project data <CircleHelp size={13} /></div>
        <div className="sidebar-user"><div className="user-avatar">{session?.name.split(' ').map((part) => part[0]).join('').slice(0, 2)}</div><div><strong>{session?.name ?? 'Teams user'}</strong><span>{roleTitle(role)}</span></div><MoreHorizontal size={16} /></div>
        {import.meta.env.DEV && <label className="role-preview">Preview role<select value={devUser} onChange={(event) => setDevUser(event.target.value)}>{demoUsers.map((user) => <option value={user.id} key={user.id}>{user.role}</option>)}</select></label>}
      </div>
    </aside>

    <main className="main-area">
      <header className="topbar"><div className="breadcrumbs"><span>Projects</span><ChevronRight size={13} /><strong>{project?.name ?? 'Your projects'}</strong>{currentTopic && <><ChevronRight size={13} /><strong>{currentTopic.name}</strong></>}</div><div className="topbar-right"><span className="sync-state"><i /> Shared and up to date</span><button className="icon-button" aria-label="Help"><CircleHelp size={17} /></button><div className="top-avatar">{session?.name.split(' ').map((part) => part[0]).join('').slice(0, 2)}</div></div></header>
      {project ? <div className="page-content">
        <section className="page-heading"><div className="eyebrow"><span /> PROJECT TRACKER <small>TEAM SPACE</small></div><div className="heading-row"><div><div className="heading-titleline"><h1>{currentTopic?.name ?? 'Project overview'}</h1>{currentTopic && canManage && <><button className="icon-button" title="Edit topic" aria-label={`Edit ${currentTopic.name}`} onClick={() => { setEditingTopic(currentTopic); setDialog('edit-topic'); }}><Pencil size={15} /></button><button className="icon-button danger-icon" title="Remove topic" aria-label={`Remove ${currentTopic.name}`} onClick={() => setPendingRemoval({ kind: 'topic', topic: currentTopic })}><Trash2 size={15} /></button></>}</div><p>{currentTopic ? `A focused view of work inside ${currentTopic.name}.` : project.description}</p></div><div className="heading-buttons"><button className="button secondary" disabled={!canEditWork} onClick={() => { setCandidates([]); setCopilotProposals(null); setCopilotPrompt(''); setTranscriptName(''); setTranscriptText(''); setDialog('transcript'); }}><FileUp size={15} /> Import transcript</button><button className="button primary" disabled={!canEditWork} onClick={() => { setEditingTask(null); setDialog('task'); }}><Plus size={16} /> Add task</button></div></div></section>
        {activeTopicId === null && <section className="overview-strip"><div className="progress-wrap"><div className="progress-ring" style={{ '--progress': `${completion}%` } as React.CSSProperties}><span>{completion}<small>%</small></span></div><div><small>PROJECT COMPLETION</small><strong>{completion > 60 ? 'Moving forward' : blocked ? 'Needs attention' : 'Getting underway'}</strong></div></div><div className="strip-divider" /><div className="metric"><small>OPEN TASKS</small><strong>{projectTasks.length - completed}<span> across {project.topics.length} topics</span></strong></div><div className="metric"><small>BLOCKED</small><strong className={blocked ? 'blocked-metric' : ''}>{String(blocked).padStart(2, '0')}<span> need follow-up</span></strong></div><div className="team-peek"><div className="avatar-stack"><i>MC</i><i>JL</i><i>SR</i></div><span>Core team</span>{canManageMembers && <button className="icon-button tiny-add" onClick={() => setDialog('members')} aria-label="Manage access"><Plus size={13} /></button>}</div></section>}
        <section className="topic-summary"><div className="summary-icon"><AlignLeft size={16} /></div><div><strong>{currentTopic ? 'Topic status' : 'Project pulse'}</strong><span>{currentTopic ? `${topicTasks.length} tasks · shared with project members` : 'A living view of work in motion'}</span></div><div className="summary-right"><div className="health-bars"><i /><i /><i /><i /><i /></div><span>{roleTitle(role)} access</span>{canManageMembers && <button className="icon-button" aria-label="Manage members" onClick={() => setDialog('members')}><UsersRound size={16} /></button>}</div></section>
        <section className="tasks-panel" style={{ '--task-font-size': `${taskFontSize}px` } as React.CSSProperties}>
          <div className="panel-heading">
            <div>
              <div className="panel-title"><h2>{showRecycleBin ? 'Recycle bin' : 'Tasks'}</h2><span>{showRecycleBin ? visibleRecycleBin.length : visibleTasks.length}</span></div>
              <p>{showRecycleBin ? 'Removed tasks stay here until restored or permanently deleted.' : 'Keep the next actions visible and moving.'}</p>
            </div>
            <div className="panel-tools">
              <label className="search-control"><Search size={14} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder={showRecycleBin ? 'Search removed tasks' : 'Search tasks'} aria-label="Search tasks" />{search && <button onClick={() => setSearch('')} aria-label="Clear search"><X size={12} /></button>}</label>
              {!showRecycleBin && <label className="font-size-control"><Type size={15} aria-hidden="true" /><span>Text size</span><input type="range" min={14} max={26} step={1} value={taskFontSize} onChange={(event) => setTaskFontSize(Number(event.target.value))} aria-label="Task text size" /><output>{taskFontSize}px</output></label>}
              {canManage && <button className="filter-button recycle-toggle" onClick={() => { setShowRecycleBin(!showRecycleBin); setFilter('All tasks'); }}><Archive size={14} />{showRecycleBin ? 'Back to tasks' : `Recycle bin (${topicRecycleBin.length})`}</button>}
              {!showRecycleBin && <button className="filter-button" onClick={() => setFilter(filter === 'All tasks' ? 'Not started' : 'All tasks')}><ListFilter size={14} />{filter === 'All tasks' ? 'Filter' : filter}</button>}
            </div>
          </div>
          {!showRecycleBin && <div className="status-tabs" role="tablist" aria-label="Filter by task status">{(['All tasks', ...statuses] as const).map((value) => <button role="tab" aria-selected={filter === value} className={filter === value ? 'selected' : ''} key={value} onClick={() => setFilter(value)}>{value}<span>{value === 'All tasks' ? topicTasks.length : topicTasks.filter((task) => task.status === value).length}</span></button>)}</div>}
          <div className="table-scroll"><table className={showRecycleBin ? 'recycle-table' : ''}><thead><tr>{showRecycleBin ? <><th>TASK</th><th>TOPIC</th><th>REMOVED</th><th>STATUS</th><th><span className="sr-only">Actions</span></th></> : <><th>TASK</th><th>OWNER</th><th>DUE DATE</th><th>STATUS</th><th><span className="sr-only">Actions</span></th></>}</tr></thead><tbody>
            {showRecycleBin ? visibleRecycleBin.map((task) => <tr key={task.id}><td><span className="deleted-task-title"><strong>{task.title}</strong><small>{task.description}</small></span></td><td>{task.topicName}</td><td>{dateLabel(task.deletedAt.slice(0, 10))}</td><td><span className="recycle-status">Removed</span></td><td><div className="recycle-actions"><button className="icon-button" title="Restore task" aria-label={`Restore ${task.title}`} onClick={() => void restoreTask(task)}><RotateCcw size={15} /></button><button className="icon-button danger-icon" title="Delete permanently" aria-label={`Delete ${task.title} permanently`} onClick={() => setPendingRemoval({ kind: 'deleted-task', task })}><Trash2 size={15} /></button></div></td></tr>) : visibleTasks.map((task) => <tr key={task.id}>
              <td>
                <div className="task-cell-wrap">
                  <button className="task-cell" onClick={() => { setEditingTask(task); setDialog('task'); }}>
                    <span className={`task-check ${task.status === 'Done' ? 'done' : ''}`}>{task.status === 'Done' && <Check size={11} />}</span>
                    <span className="task-copy"><strong className="task-title">{task.title}</strong><span className="task-description">{task.description || 'No description provided.'}</span></span>
                  </button>
                </div>
                <section className="task-comments-section" aria-label={`Comments for ${task.title}`}>
                  <div className="task-comments-heading">Comments <span>{task.comments.length}</span></div>
                  {task.comments.length ? <div className="task-comments-list">{task.comments.map((comment, index) => <p key={`${task.id}-${index}`}>{comment}</p>)}</div> : <p className="no-task-comments">No comments yet.</p>}
                  {canEditWork && task.status !== 'Done' && <div className="task-comment-compose"><textarea value={commentDrafts[task.id] ?? ''} onChange={(event) => setCommentDrafts((drafts) => ({ ...drafts, [task.id]: event.target.value }))} placeholder="Add a comment..." rows={1} aria-label={`New comment for ${task.title}`} /><button className="button secondary" disabled={!commentDrafts[task.id]?.trim()} onClick={() => void saveTaskComment(task)}>Add</button></div>}
                </section>
              </td>
              <td><span className="owner-initials">{task.owner ? task.owner.split(' ').map((part) => part[0]).join('').slice(0, 2).toUpperCase() : <UsersRound size={13} />}</span><span className="owner-name">{task.owner || 'Unassigned'}</span></td>
              <td><span className="date-value"><CalendarDays size={13} />{dateLabel(task.dueDate)}</span></td>
              <td><select disabled={!canEditWork || task.status === 'Done'} className={`status-pill ${statusKey(task.status)}`} value={task.status} onChange={(event) => void updateStatus(task, event.target.value as Status)} aria-label={`Status for ${task.title}`}>{statuses.map((status) => <option key={status}>{status}</option>)}</select></td>
              <td><button className="icon-button row-edit" aria-label={`Open ${task.title}`} onClick={() => { setEditingTask(task); setDialog('task'); }}><MoreHorizontal size={16} /></button></td>
            </tr>)}
            {(showRecycleBin ? !visibleRecycleBin.length : !visibleTasks.length) && <tr><td colSpan={5}><div className="empty-list"><div>{showRecycleBin ? <Archive size={18} /> : <Search size={18} />}</div><strong>{search ? 'No matching tasks' : showRecycleBin ? 'Recycle bin is empty' : 'Nothing on this list yet'}</strong><span>{search ? 'Try another search.' : showRecycleBin ? 'Removed tasks for this topic will appear here.' : 'Add a task or import a transcript to get started.'}</span>{!showRecycleBin && !search && canEditWork && <button className="button primary" onClick={() => { setEditingTask(null); setDialog('task'); }}><Plus size={15} /> Add first task</button>}</div></td></tr>}
          </tbody></table></div><div className="table-foot"><span>{showRecycleBin ? `${visibleRecycleBin.length} of ${topicRecycleBin.length} removed tasks` : `Showing ${visibleTasks.length} of ${topicTasks.length} tasks`}</span>{!showRecycleBin && canEditWork && <button onClick={() => { setEditingTask(null); setDialog('task'); }}><Plus size={13} /> Add a task</button>}</div>
        </section>
        <footer className="page-footer"><span>FIELDNOTES <i>·</i> PROJECT WORKSPACE</span><span>Clear work, shared progress.</span></footer>
      </div> : <div className="no-project"><div className="empty-project-icon"><FolderKanban size={23} /></div><span className="eyebrow">FIELDNOTES WORKSPACE</span><h1>Start with a shared project.</h1><p>Create a project and invite teammates with the role that fits their work.</p><button className="button primary" onClick={() => setDialog('project')}><Plus size={16} /> Create project</button></div>}
    </main>

    {dialog === 'project' && <ProjectDialog onClose={() => setDialog(null)} onSave={createProject} />}
    {dialog === 'edit-project' && project && <ProjectDialog initial={project} onClose={() => setDialog(null)} onSave={updateProject} />}
    {dialog === 'topic' && project && <TopicDialog topics={project.topics} selectedParentId={activeTopicId} onClose={() => setDialog(null)} onSave={addTopic} />}
    {dialog === 'edit-topic' && project && editingTopic && <TopicDialog topics={project.topics} selectedParentId={editingTopic.parentId} initialTopic={editingTopic} onClose={() => setDialog(null)} onSave={(name, parentId) => updateTopic(editingTopic, name, parentId)} />}
    {dialog === 'task' && project && <TaskDialog project={project} task={editingTask} role={role} defaultTopicId={activeTopicId ?? project.topics[0]?.id ?? ''} onClose={() => setDialog(null)} onCreate={addTask} onSave={updateTask} onRemove={(task) => { setDialog(null); setEditingTask(null); setPendingRemoval({ kind: 'task', task }); }} />}
    {dialog === 'transcript' && project && <TranscriptDialog project={project} transcriptName={transcriptName} candidates={candidates} copilotProposals={copilotProposals} text={transcriptText} copilotPrompt={copilotPrompt} busy={busy} onClose={() => setDialog(null)} onCandidates={setCandidates} onCopilotProposals={setCopilotProposals} onText={setTranscriptText} onName={setTranscriptName} onAnalyzeText={analyzeTranscriptResult} onImport={importCandidates} onCopilot={prepareWithCopilot} onCopyCopilotPrompt={copyCopilotPrompt} onPasteCopilotResult={pasteCopilotResult} />}
    {dialog === 'members' && project && canManageMembers && <MembersDialog projectId={project.id} role={role} devUser={devUser} token={token} onClose={() => setDialog(null)} onChanged={() => void loadWorkspace(devUser, token)} />}
    {pendingRemoval && <RemovalDialog target={pendingRemoval} onCancel={() => setPendingRemoval(null)} onConfirm={() => void confirmRemoval()} />}
    {(notice || error) && <div className={`toast ${error ? 'toast-error' : ''}`} role="status">{error ? <CircleHelp size={15} /> : <Check size={15} />}{error || notice}<button onClick={() => { setNotice(''); setError(''); }} aria-label="Dismiss"><X size={14} /></button></div>}
  </div>;
}

function TopicBranch({ topic, topics, tasks, activeId, expanded, canManage, onSelect, onExpand, onAddChild, onEdit, onRemove, onMove, depth = 0 }: {
  topic: Topic; topics: Topic[]; tasks: Task[]; activeId: string | null; expanded: string[]; canManage: boolean;
  onSelect: (id: string) => void; onExpand: (id: string) => void; onAddChild: (id: string) => void; onEdit: (topic: Topic) => void; onRemove: (topic: Topic) => void; onMove: (id: string, direction: -1 | 1) => void; depth?: number;
}) {
  const children = topics.filter((item) => item.parentId === topic.id);
  const siblings = topics.filter((item) => item.parentId === topic.parentId);
  const isExpanded = expanded.includes(topic.id);
  const count = tasks.filter((task) => task.topicId === topic.id).length;
  return <div className="topic-branch"><div className={`tree-row ${activeId === topic.id ? 'active' : ''}`} style={{ '--depth': depth } as React.CSSProperties}>
    <button className="tree-expand" onClick={() => onExpand(topic.id)} aria-label={isExpanded ? 'Collapse topic' : 'Expand topic'}>{children.length ? isExpanded ? <ChevronDown size={13} /> : <ChevronRight size={13} /> : <i />}</button>
    <button className="tree-topic" onClick={() => onSelect(topic.id)}><i className="topic-diamond" /><span>{topic.name}</span><b>{count}</b></button>
    {canManage && <div className="tree-actions"><button onClick={() => onAddChild(topic.id)} aria-label={`Add subtopic to ${topic.name}`} title="Add subtopic"><Plus size={12} /></button><button onClick={() => onEdit(topic)} aria-label={`Edit ${topic.name}`} title="Edit topic"><Pencil size={12} /></button><button disabled={siblings[0]?.id === topic.id} onClick={() => onMove(topic.id, -1)} aria-label={`Move ${topic.name} up`} title="Move up"><ArrowUp size={12} /></button><button disabled={siblings.at(-1)?.id === topic.id} onClick={() => onMove(topic.id, 1)} aria-label={`Move ${topic.name} down`} title="Move down"><ArrowDown size={12} /></button><button className="danger-icon" onClick={() => onRemove(topic)} aria-label={`Remove ${topic.name}`} title="Remove topic"><Trash2 size={12} /></button></div>}
  </div>{children.length > 0 && isExpanded && children.map((child) => <TopicBranch key={child.id} topic={child} topics={topics} tasks={tasks} activeId={activeId} expanded={expanded} canManage={canManage} onSelect={onSelect} onExpand={onExpand} onAddChild={onAddChild} onEdit={onEdit} onRemove={onRemove} onMove={onMove} depth={depth + 1} />)}</div>;
}

function Modal({ title, eyebrow, onClose, children, className = '' }: { title: string; eyebrow: string; onClose: () => void; children: React.ReactNode; className?: string }) {
  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className={`modal ${className}`} role="dialog" aria-modal="true" aria-label={title}><div className="modal-heading"><span>{eyebrow}</span><button className="icon-button" onClick={onClose} aria-label="Close"><X size={17} /></button><h2>{title}</h2></div>{children}</section></div>;
}

function ProjectDialog({ initial, onClose, onSave }: { initial?: Pick<Project, 'name' | 'description'>; onClose: () => void; onSave: (name: string, description: string) => Promise<void> }) {
  const [name, setName] = useState(initial?.name ?? '');
  const [description, setDescription] = useState(initial?.description ?? '');
  return <Modal title={initial ? 'Edit project' : 'Create a project'} eyebrow={initial ? 'PROJECT SETTINGS' : 'NEW PROJECT'} onClose={onClose}><form className="modal-form" onSubmit={(event) => { event.preventDefault(); if (name.trim()) void onSave(name.trim(), description.trim()); }}><label>Project name<input autoFocus required value={name} onChange={(event) => setName(event.target.value)} placeholder="e.g. Website refresh" /></label><label>Project outcome<textarea rows={3} value={description} onChange={(event) => setDescription(event.target.value)} placeholder="What are you working toward?" /></label><div className="modal-actions"><button type="button" className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" type="submit">{initial ? <Check size={15} /> : <Plus size={15} />}{initial ? 'Save changes' : 'Create project'}</button></div></form></Modal>;
}

function getTopicDescendants(topics: Topic[], rootId: string) {
  const descendants = new Set([rootId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const topic of topics) {
      if (topic.parentId && descendants.has(topic.parentId) && !descendants.has(topic.id)) {
        descendants.add(topic.id);
        changed = true;
      }
    }
  }
  return descendants;
}

function TopicDialog({ topics, selectedParentId, initialTopic, onClose, onSave }: { topics: Topic[]; selectedParentId: string | null; initialTopic?: Topic; onClose: () => void; onSave: (name: string, parentId: string | null) => Promise<void> }) {
  const [name, setName] = useState(initialTopic?.name ?? '');
  const [parentId, setParentId] = useState(initialTopic?.parentId ?? selectedParentId ?? '');
  const excludedIds = initialTopic ? getTopicDescendants(topics, initialTopic.id) : new Set<string>();
  const parentTopics = topics.filter((topic) => !excludedIds.has(topic.id));
  return <Modal title={initialTopic ? 'Edit topic' : 'Add a topic'} eyebrow="PROJECT STRUCTURE" onClose={onClose}><form className="modal-form" onSubmit={(event) => { event.preventDefault(); if (name.trim()) void onSave(name.trim(), parentId || null); }}><label>Topic name<input autoFocus required value={name} onChange={(event) => setName(event.target.value)} placeholder="e.g. Customer research" /></label><label>Place under<select value={parentId} onChange={(event) => setParentId(event.target.value)}><option value="">Project root</option>{parentTopics.map((topic) => <option key={topic.id} value={topic.id}>{topic.name}</option>)}</select></label><p className="field-note">Topics can be nested and reordered from the project sidebar.</p><div className="modal-actions"><button type="button" className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" type="submit">{initialTopic ? <Check size={15} /> : <Plus size={15} />}{initialTopic ? 'Save changes' : 'Add topic'}</button></div></form></Modal>;
}

function TaskDialog({ project, task, role, defaultTopicId, onClose, onCreate, onSave, onRemove }: { project: Project; task: Task | null; role: Role; defaultTopicId: string; onClose: () => void; onCreate: (task: Omit<Task, 'id'>) => Promise<void>; onSave: (task: Task, changes: Partial<Task>) => Promise<void>; onRemove: (task: Task) => void }) {
  const [title, setTitle] = useState(task?.title ?? '');
  const [description, setDescription] = useState(task?.description ?? '');
  const [dueDate, setDueDate] = useState(task?.dueDate ?? '');
  const [status, setStatus] = useState<Status>(task?.status ?? 'Not started');
  const [owner, setOwner] = useState(task?.owner ?? '');
  const [topicId, setTopicId] = useState(task?.topicId ?? defaultTopicId);
  const [comments, setComments] = useState(task?.comments ?? []);
  const [comment, setComment] = useState('');
  const canEdit = role === 'owner' || role === 'manager';
  const isCompleted = task?.status === 'Done';
  const canEditDetails = canEdit && !isCompleted;
  const canChangeStatus = role !== 'viewer' && !isCompleted;
  const canComment = role !== 'viewer' && !isCompleted;
  return <Modal title={task ? 'Task details' : 'Add a task'} eyebrow={task ? 'ACTION ITEM' : 'NEW ACTION'} onClose={onClose} className="task-modal"><form className="modal-form" onSubmit={(event) => { event.preventDefault(); if (!title.trim()) return; if (task) { const changes = canEdit ? { title: title.trim(), description: description.trim(), dueDate, status, owner: owner.trim(), topicId, comments } : { status, comments }; void onSave(task, changes); } else { void onCreate({ topicId, title: title.trim(), description: description.trim(), dueDate, status, owner: owner.trim(), comments: [] }); } }}>
    <label>Title<input autoFocus required disabled={Boolean(task) && !canEditDetails} value={title} onChange={(event) => setTitle(event.target.value)} placeholder="What needs to happen?" /></label>
    <label>Description<textarea disabled={Boolean(task) && !canEditDetails} rows={3} value={description} onChange={(event) => setDescription(event.target.value)} placeholder="Add context, criteria, or useful links." /></label>
    <div className="form-grid"><label>Due date<input type="date" disabled={Boolean(task) && !canEditDetails} value={dueDate} onChange={(event) => setDueDate(event.target.value)} /></label><label>Status<select disabled={!canChangeStatus} value={status} onChange={(event) => setStatus(event.target.value as Status)}>{statuses.map((value) => <option key={value}>{value}</option>)}</select></label></div>
    <div className="form-grid"><label>Owner<input disabled={Boolean(task) && !canEditDetails} value={owner} onChange={(event) => setOwner(event.target.value)} placeholder="Name" /></label><label>Topic<select disabled={Boolean(task) && !canEditDetails} value={topicId} onChange={(event) => setTopicId(event.target.value)}>{project.topics.map((topic) => <option value={topic.id} key={topic.id}>{topic.name}</option>)}</select></label></div>
    {task && <div className="comments-block"><div className="comments-title"><MessageSquareText size={14} /> Comments <span>{comments.length}</span></div>{comments.map((value, index) => <p key={`${index}-${value}`} className="comment-item">{value}</p>)}{canComment && <div className="comment-compose"><input value={comment} onChange={(event) => setComment(event.target.value)} placeholder="Write a comment..." /><button type="button" className="button secondary" disabled={!comment.trim()} onClick={() => { setComments([...comments, comment.trim()]); setComment(''); }}>Add</button></div>}</div>}
    <div className="modal-actions modal-actions-split">{task && canEdit && <button type="button" className="button danger-button" onClick={() => onRemove(task)}><Trash2 size={15} /> Remove task</button>}<span /> <button type="button" className="button secondary" onClick={onClose}>Close</button>{(!task || role !== 'viewer') && !isCompleted && <button className="button primary" type="submit"><Check size={15} /> {task ? 'Save changes' : 'Create task'}</button>}</div>
  </form></Modal>;
}

function RemovalDialog({ target, onCancel, onConfirm }: { target: RemovalTarget; onCancel: () => void; onConfirm: () => void }) {
  const isProject = target.kind === 'project';
  const isDeletedTask = target.kind === 'deleted-task';
  const title = isProject ? target.project.name : target.kind === 'topic' ? target.topic.name : target.task.title;
  const heading = isProject ? 'Remove project?' : target.kind === 'topic' ? 'Remove topic?' : isDeletedTask ? 'Delete task permanently?' : 'Move task to recycle bin?';
  const detail = isProject
    ? 'This permanently removes the project, its topics, tasks, recycle-bin items, and project access.'
    : target.kind === 'topic'
      ? 'This removes the topic and its subtopics. Their tasks will be moved to the project recycle bin.'
      : isDeletedTask
        ? 'This task will be permanently deleted and cannot be restored.'
        : 'The task will be removed from the topic and can be restored from its recycle bin.';
  const action = isProject ? 'Remove project' : target.kind === 'topic' ? 'Remove topic' : isDeletedTask ? 'Delete permanently' : 'Move to recycle bin';
  return <Modal title={heading} eyebrow="CONFIRM REMOVAL" onClose={onCancel} className="removal-modal"><div className="removal-copy"><strong>{title}</strong><p>{detail}</p></div><div className="modal-actions"><button className="button secondary" onClick={onCancel}>Cancel</button><button className="button danger-button" onClick={onConfirm}><Trash2 size={15} />{action}</button></div></Modal>;
}

function TranscriptDialog({ project, transcriptName, candidates, copilotProposals, text, copilotPrompt, busy, onClose, onCandidates, onCopilotProposals, onText, onName, onAnalyzeText, onImport, onCopilot, onCopyCopilotPrompt, onPasteCopilotResult }: { project: Project; transcriptName: string; candidates: Candidate[]; copilotProposals: CopilotProposal[] | null; text: string; copilotPrompt: string; busy: boolean; onClose: () => void; onCandidates: (items: Candidate[]) => void; onCopilotProposals: (items: CopilotProposal[] | null) => void; onText: (value: string) => void; onName: (value: string) => void; onAnalyzeText: (value: string) => void; onImport: () => Promise<void>; onCopilot: () => Promise<void>; onCopyCopilotPrompt: () => Promise<void>; onPasteCopilotResult: () => void }) {
  const [pasteMode, setPasteMode] = useState(false);
  const [fileError, setFileError] = useState('');
  const selectedCount = candidates.filter((candidate) => candidate.selected).length;
  async function readFile(file?: File) {
    if (!file) return;
    setFileError('');
    if (file.size > 5 * 1024 * 1024) {
      setFileError('Choose a file smaller than 5 MB.');
      return;
    }
    try {
      const extension = file.name.split('.').pop()?.toLowerCase();
      let value: string;
      if (extension === 'docx') {
        const { default: mammoth } = await import('mammoth');
        value = (await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() })).value;
      } else if (extension === 'vtt') {
        value = stripVttMetadata(await file.text());
      } else {
        value = await file.text();
      }
      const extracted = parseTranscript(value);
      onCopilotProposals(null);
      onName(file.name);
      onText(value);
      onCandidates(extracted);
      if (!extracted.length) setFileError('No likely action items found. Look for lines with “TODO”, “will”, or “needs to”.');
    } catch {
      setFileError('This file could not be read. Try a DOCX, WebVTT, TXT, or Markdown file.');
    }
  }
  return <Modal title="Turn a conversation into action" eyebrow="TRANSCRIPT IMPORT" onClose={onClose} className="transcript-modal">
    {copilotPrompt && <section className="copilot-prompt-panel" aria-label="Microsoft 365 Copilot handoff">
      <div><strong>Prompt ready for Microsoft 365 Copilot</strong><p>Copy the prompt, paste it into Copilot, then paste Copilot’s task lines here for review. The transcript is not sent automatically.</p></div>
      <textarea readOnly value={copilotPrompt} rows={5} aria-label="Prompt to paste into Copilot" />
      <div className="copilot-prompt-actions"><button className="button secondary" onClick={() => void onCopyCopilotPrompt()}><Copy size={14} /> Copy prompt</button><a className="button secondary" href="https://m365.cloud.microsoft/chat" target="_blank" rel="noopener noreferrer">Open Microsoft 365 Copilot <ExternalLink size={14} /></a><button className="button primary" onClick={() => { onPasteCopilotResult(); setPasteMode(true); }}>Paste Copilot result</button></div>
    </section>}
    {!candidates.length && copilotProposals === null ? <div className="transcript-empty"><p>Choose meeting notes or paste a transcript. We’ll suggest action items for you to review before they’re added.</p>
      {!pasteMode ? <><label className="dropzone"><input type="file" accept=".txt,.md,.vtt,.docx,text/plain,text/markdown,text/vtt,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={(event) => { void readFile(event.target.files?.[0]); event.target.value = ''; }} /><span className="upload-icon"><FileUp size={19} /></span><strong>Choose a transcript file</strong><small>TXT, Markdown, DOCX, or WebVTT · up to 5 MB</small><em>Browse files</em></label>{fileError && <p className="inline-error">{fileError}</p>}{text.trim() && <button className="text-link copilot-handoff-link" onClick={() => void onCopilot()}>Prepare transcript with Microsoft 365 Copilot</button>}<button className="text-link" onClick={() => { setFileError(''); setPasteMode(true); }}>Or paste transcript text</button></> : <><textarea autoFocus className="transcript-textarea" value={text} onChange={(event) => onText(event.target.value)} placeholder="Paste transcript or Copilot JSON response here..." rows={8} /><div className="modal-actions"><button className="button secondary" onClick={() => setPasteMode(false)}>Back</button>{text.trim() && <button className="button secondary" onClick={() => void onCopilot()}>Prepare with Copilot <Sparkles size={14} /></button>}<button className="button primary" disabled={!text.trim()} onClick={() => onAnalyzeText(text)}>Review suggestions <Sparkles size={14} /></button></div></>}
      <div className="privacy-note"><Check size={13} /> Transcript stays in this Teams workspace.</div>
    </div> : copilotProposals !== null ? <CopilotProposalReview project={project} proposals={copilotProposals} busy={busy} onProposals={onCopilotProposals} onClose={onClose} onImport={onImport} onStartOver={() => { onCopilotProposals(null); onCandidates([]); onText(''); onName(''); }} /> : <div className="candidate-content"><div className="candidate-head"><div><strong>{candidates.length} possible action items</strong><span>From {transcriptName || 'transcript'} · review before adding</span></div><div className="candidate-actions"><button className="text-link copilot-handoff-link" onClick={() => void onCopilot()}>Prepare with Microsoft 365 Copilot</button><button className="text-link" onClick={() => { onCandidates([]); onText(''); onName(''); }}>Start over</button></div></div><div className="candidate-list">{candidates.map((candidate, index) => <article className="candidate-card" key={`${index}-${candidate.title}`}><label className="candidate-check"><input type="checkbox" checked={candidate.selected} onChange={(event) => onCandidates(candidates.map((item, position) => position === index ? { ...item, selected: event.target.checked } : item))} /><i><Check size={11} /></i></label><div className="candidate-fields"><input className="candidate-title" value={candidate.title} aria-label="Task title" onChange={(event) => onCandidates(candidates.map((item, position) => position === index ? { ...item, title: event.target.value } : item))} /><div className="candidate-meta"><label><UsersRound size={12} /><input value={candidate.owner} placeholder="Owner" onChange={(event) => onCandidates(candidates.map((item, position) => position === index ? { ...item, owner: event.target.value } : item))} /></label><label><CalendarDays size={12} /><input type="date" value={candidate.dueDate} onChange={(event) => onCandidates(candidates.map((item, position) => position === index ? { ...item, dueDate: event.target.value } : item))} /></label><select value={candidate.status} onChange={(event) => onCandidates(candidates.map((item, position) => position === index ? { ...item, status: event.target.value as Status } : item))}>{statuses.map((status) => <option key={status}>{status}</option>)}</select></div><p>{candidate.description}</p></div></article>)}</div><div className="modal-actions"><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!selectedCount || busy} onClick={() => void onImport()}><Check size={15} /> Add {selectedCount} task{selectedCount === 1 ? '' : 's'}</button></div></div>}
  </Modal>;
}

function CopilotProposalReview({ project, proposals, busy, onProposals, onClose, onImport, onStartOver }: {
  project: Project;
  proposals: CopilotProposal[];
  busy: boolean;
  onProposals: (proposals: CopilotProposal[]) => void;
  onClose: () => void;
  onImport: () => Promise<void>;
  onStartOver: () => void;
}) {
  const selectedCount = proposals.filter((proposal) => proposal.selected).length;
  const changeLabels: Record<string, string> = {
    title: 'Title',
    description: 'Description',
    dueDate: 'Due date',
    status: 'Status',
    owner: 'Owner',
  };

  return <div className="candidate-content">
    <div className="candidate-head">
      <div><strong>{proposals.length} proposed project changes</strong><span>Review each suggestion before applying it.</span></div>
      <button className="text-link" onClick={onStartOver}>Start over</button>
    </div>
    {proposals.length ? <div className="candidate-list">
      {proposals.map((proposal, index) => <article className="copilot-proposal" key={proposal.action === 'update' ? proposal.taskId : `${proposal.topicId}-${proposal.title}-${index}`}>
        <label className="candidate-check"><input type="checkbox" checked={proposal.selected} onChange={(event) => onProposals(proposals.map((item, position) => position === index ? { ...item, selected: event.target.checked } : item))} /><i><Check size={11} /></i></label>
        <div className="copilot-proposal-content">
          <span className={`proposal-kind ${proposal.action}`}>{proposal.action === 'create' ? 'NEW TASK' : 'UPDATE TASK'}</span>
          {proposal.action === 'create' ? <>
            <strong>{proposal.title}</strong>
            <span className="proposal-topic">Topic: {project.topics.find((topic) => topic.id === proposal.topicId)?.name ?? 'Unknown topic'}</span>
            <dl className="proposal-changes">
              <div><dt>Status</dt><dd>{proposal.status}</dd></div>
              <div><dt>Owner</dt><dd>{proposal.owner || 'Unassigned'}</dd></div>
              {proposal.dueDate && <div><dt>Due date</dt><dd>{proposal.dueDate}</dd></div>}
              {proposal.description && <div><dt>Description</dt><dd>{proposal.description}</dd></div>}
            </dl>
          </> : <>
            <strong>{proposal.taskTitle}</strong>
            <dl className="proposal-changes">{Object.entries(proposal.changes).map(([field, value]) => <div key={field}><dt>{changeLabels[field] ?? field}</dt><dd>{String(value)}</dd></div>)}</dl>
          </>}
        </div>
      </article>)}
    </div> : <div className="empty-proposals"><strong>No changes proposed</strong><span>Copilot found no clear new tasks or updates in the transcript.</span></div>}
    <div className="modal-actions"><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!selectedCount || busy} onClick={() => void onImport()}><Check size={15} /> Apply {selectedCount} change{selectedCount === 1 ? '' : 's'}</button></div>
  </div>;
}

function MembersDialog({ projectId, role, devUser, token, onClose, onChanged }: { projectId: string; role: Role; devUser: string; token: string; onClose: () => void; onChanged: () => void }) {
  const [members, setMembers] = useState<Member[]>([]);
  const [email, setEmail] = useState('');
  const [memberRole, setMemberRole] = useState<Role>('contributor');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  async function load() {
    try { const result = await request<{ members: Member[] }>(`/api/projects/${projectId}/members`, devUser, token); setMembers(result.members); }
    catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not load members.'); }
  }
  useEffect(() => { void load(); }, [projectId]);
  async function addMember(event: React.FormEvent) {
    event.preventDefault();
    setBusy(true);
    try { await request(`/api/projects/${projectId}/members`, devUser, token, { method: 'POST', body: JSON.stringify({ email, role: memberRole }) }); setEmail(''); await load(); onChanged(); }
    catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not add member.'); }
    finally { setBusy(false); }
  }
  async function changeRole(member: Member, nextRole: Role) {
    try { await request(`/api/projects/${projectId}/members/${encodeURIComponent(member.id)}`, devUser, token, { method: 'PATCH', body: JSON.stringify({ role: nextRole }) }); await load(); onChanged(); }
    catch (reason) { setError(reason instanceof Error ? reason.message : 'Could not update role.'); }
  }
  return <Modal title="Project access" eyebrow="TEAM MEMBERS" onClose={onClose} className="members-modal"><div className="members-list">{members.map((member) => <div className="member-row" key={member.id}><span className="member-avatar">{member.name.split(/[ .@]/).filter(Boolean).map((part) => part[0]).join('').slice(0, 2).toUpperCase()}</span><div className="member-copy"><strong>{member.name}</strong><span>{member.email}</span></div>{role === 'owner' && member.role !== 'owner' ? <select value={member.role} onChange={(event) => void changeRole(member, event.target.value as Role)} aria-label={`Role for ${member.name}`}><option value="owner">Owner</option><option value="manager">Manager</option><option value="contributor">Contributor</option><option value="viewer">Viewer</option></select> : <span className={`role-tag ${member.role}`}>{roleTitle(member.role)}</span>}</div>)}</div>{role !== 'viewer' && <form className="invite-form" onSubmit={(event) => void addMember(event)}><label>Invite by email<input type="email" required value={email} onChange={(event) => setEmail(event.target.value)} placeholder="name@company.com" /></label><label>Project role<select value={memberRole} onChange={(event) => setMemberRole(event.target.value as Role)}><option value="contributor">Contributor</option><option value="viewer">Viewer</option>{role === 'owner' && <><option value="manager">Manager</option><option value="owner">Owner</option></>}</select></label><button className="button primary" disabled={busy}><Plus size={14} /> Add member</button></form>}{error && <p className="inline-error">{error}</p>}<p className="role-note">Owners manage access. Managers organize project work. Contributors create tasks and update status or comments. Viewers have read-only access.</p></Modal>;
}