---
name: Teams Facilitator Action Items
description: "Use during or after a Teams meeting to have Facilitator create Fieldnotes-importable action items from the meeting transcript and notes."
argument-hint: "Use with the current Teams meeting context; optionally specify the target project or topic."
agent: "agent"
---
Use the transcript and meeting notes available in the current Teams meeting to create a plain-text action-item record for import into the Fieldnotes project workspace. Do not ask the user to paste a transcript. Do not claim to provide a verbatim transcript; this output is the actionable record extracted from the meeting.

Use any project or topic named in the request only to guide task wording. The user selects the destination project and topic in Fieldnotes before importing.

Return only actionable items, one per line, using this exact structure:

`TODO: <concise action title> [by YYYY-MM-DD] assigned to <Full Name or Unassigned> | status <Not started|In progress|Blocked|Done> | Details: <brief context>`

Example:

`TODO: Send the launch checklist by 2026-11-18 assigned to Maya Chen | status In progress | Details: Include the final review links and confirm the checklist with the team.`

Rules:
- Start every action line with `TODO:`. Use a short, imperative task title.
- Include a due date only when the source gives one. Convert explicit dates to `YYYY-MM-DD`; do not infer dates.
- Use the named owner only when the source identifies one. Otherwise write `Unassigned`.
- Use `Blocked`, `In progress`, or `Done` only when clearly stated. Otherwise use `Not started`.
- Put owner and status metadata after the title and due date, in the order shown. Keep the `|` separators; they let the importer distinguish the owner from status and details. Put supporting context after `| Details:` so it stays out of the task title.
- Keep each item on a single line. Do not include transcript dialogue, headings, bullets, summaries, code fences, or commentary; ordinary dialogue can be mistaken for tasks by the importer.
- Merge duplicate action items. Do not invent tasks, owners, dates, status changes, or decisions.
- The importer stores the full action line as task context; comments are added in the app after import.
- If no transcript or meeting notes are available, say so plainly instead of inventing action items.
- If there are no actionable tasks, return exactly: `No actionable tasks found.`