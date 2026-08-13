import express from 'express';
import cors from 'cors';

const app = express();
app.use(cors({ origin: true }));
app.use(express.json());

const port = process.env.PORT || 4000;

const formatDateValue = (date) => {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

export const extractRelativeDateFromText = (text, now = new Date()) => {
  if (!text || typeof text !== 'string') return null;

  const match = text.match(/(?:older|before|earlier|over|more than)\s+(?:than\s+)?(\d+)\s*(day|days|d|week|weeks|w|month|months|mo|year|years|y)/i);
  if (!match) return null;

  const amount = Number(match[1]);
  const unit = match[2].toLowerCase();
  const target = new Date(now);

  if (['day', 'days', 'd'].includes(unit)) {
    target.setUTCDate(target.getUTCDate() - amount);
  } else if (['week', 'weeks', 'w'].includes(unit)) {
    target.setUTCDate(target.getUTCDate() - amount * 7);
  } else if (['month', 'months', 'mo'].includes(unit)) {
    target.setUTCMonth(target.getUTCMonth() - amount);
  } else if (['year', 'years', 'y'].includes(unit)) {
    target.setUTCFullYear(target.getUTCFullYear() - amount);
  } else {
    return null;
  }

  return formatDateValue(target);
};

export const normalizeCleanupRule = (rule) => {
  if (!rule || typeof rule !== 'object') return rule;

  const nextRule = { ...rule };
  const relativeDate = extractRelativeDateFromText(nextRule.explanation || '');

  if (relativeDate && (!nextRule.receivedBefore || String(nextRule.receivedBefore).toLowerCase() === 'none')) {
    nextRule.receivedBefore = relativeDate;
  }

  if (nextRule.receivedBefore === 'none' || nextRule.receivedBefore === 'null') {
    nextRule.receivedBefore = '';
  }

  if (typeof nextRule.hasAttachments === 'boolean') {
    nextRule.hasAttachments = nextRule.hasAttachments ? 'yes' : 'no';
  } else if (typeof nextRule.hasAttachments === 'string') {
    const attachmentsValue = nextRule.hasAttachments.trim().toLowerCase();
    if (attachmentsValue === 'true') nextRule.hasAttachments = 'yes';
    else if (attachmentsValue === 'false') nextRule.hasAttachments = 'no';
    else if (!['yes', 'no', 'any', ''].includes(attachmentsValue)) nextRule.hasAttachments = '';
    else nextRule.hasAttachments = attachmentsValue;
  }

  if (typeof nextRule.readStatus === 'string') {
    const statusValue = nextRule.readStatus.trim().toLowerCase();
    if (['read', 'unread', 'any'].includes(statusValue)) {
      nextRule.readStatus = statusValue;
    } else {
      nextRule.readStatus = 'any';
    }
  }

  if (typeof nextRule.importance === 'string') {
    const importanceValue = nextRule.importance.trim().toLowerCase();
    if (['low', 'normal', 'high', 'any'].includes(importanceValue)) {
      nextRule.importance = importanceValue;
    } else {
      nextRule.importance = 'any';
    }
  }

  if (typeof nextRule.folder === 'string' && nextRule.folder.toLowerCase() === 'none') {
    nextRule.folder = '';
  }

  if (typeof nextRule.body === 'string' && nextRule.body.toLowerCase() === 'none') {
    nextRule.body = '';
  }

  if (typeof nextRule.receivedAfter === 'string' && (nextRule.receivedAfter.toLowerCase() === 'none' || nextRule.receivedAfter.toLowerCase() === 'null')) {
    nextRule.receivedAfter = '';
  }

  return nextRule;
};

const jobs = new Map();

const createJobId = () => `job-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

export const buildGraphFilter = ({ sender = '', subject = '', body = '', folder = '', receivedAfter = '', receivedBefore = '', readStatus = 'any', hasAttachments = 'any', importance = 'any', senderExclude = '', subjectExclude = '' }) => {
  const filterParts = [];

  if (sender.trim()) {
    const senders = sender.split(',').map((s) => s.trim()).filter(Boolean);
    if (senders.length === 1) {
      const safeSender = senders[0].replace(/'/g, "''");
      filterParts.push(`contains(from/emailAddress/address,'${safeSender}')`);
    } else if (senders.length > 1) {
      const orParts = senders.map((s) => {
        const safe = s.replace(/'/g, "''");
        return `contains(from/emailAddress/address,'${safe}')`;
      });
      filterParts.push(`(${orParts.join(' or ')})`);
    }
  }

  if (senderExclude.trim()) {
    const senders = senderExclude.split(',').map((s) => s.trim()).filter(Boolean);
    senders.forEach((s) => {
      const safe = s.replace(/'/g, "''");
      filterParts.push(`not(contains(from/emailAddress/address,'${safe}'))`);
    });
  }

  if (subject.trim()) {
    const safeSubject = subject.trim().replace(/'/g, "''");
    filterParts.push(`contains(subject,'${safeSubject}')`);
  }

  if (subjectExclude.trim()) {
    const safeSubject = subjectExclude.trim().replace(/'/g, "''");
    filterParts.push(`not(contains(subject,'${safeSubject}'))`);
  }

  if (body.trim()) {
    const safeBody = body.trim().replace(/'/g, "''");
    filterParts.push(`contains(body/content,'${safeBody}')`);
  }

  if (folder.trim()) {
    const safeFolder = folder.trim().replace(/'/g, "''");
    filterParts.push(`parentFolderId eq '${safeFolder}'`);
  }

  if (receivedAfter.trim()) {
    const afterDate = new Date(receivedAfter);
    afterDate.setUTCHours(0, 0, 0, 0);
    filterParts.push(`receivedDateTime ge ${afterDate.toISOString()}`);
  }

  if (receivedBefore.trim()) {
    const beforeDate = new Date(receivedBefore);
    beforeDate.setUTCHours(23, 59, 59, 999);
    filterParts.push(`receivedDateTime le ${beforeDate.toISOString()}`);
  }

  if (readStatus === 'read') {
    filterParts.push('isRead eq true');
  } else if (readStatus === 'unread') {
    filterParts.push('isRead eq false');
  }

  if (hasAttachments === 'yes') {
    filterParts.push('hasAttachments eq true');
  } else if (hasAttachments === 'no') {
    filterParts.push('hasAttachments eq false');
  }

  if (importance !== 'any') {
    const safeImportance = importance.replace(/'/g, "''");
    filterParts.push(`importance eq '${safeImportance}'`);
  }

  return filterParts.length > 0 ? filterParts.join(' and ') : '';
};

const summarizeCriteria = (criteria) => {
  const parts = [];
  if (criteria.sender) parts.push(`sender contains '${criteria.sender}'`);
  if (criteria.senderExclude) parts.push(`sender not contains '${criteria.senderExclude}'`);
  if (criteria.subject) parts.push(`subject contains '${criteria.subject}'`);
  if (criteria.subjectExclude) parts.push(`subject not contains '${criteria.subjectExclude}'`);
  if (criteria.body) parts.push(`body contains '${criteria.body}'`);
  const folderLabel = criteria.folderName || criteria.folder || '';
  if (folderLabel) parts.push(`folder '${folderLabel}'`);
  if (criteria.receivedAfter) parts.push(`after ${criteria.receivedAfter}`);
  if (criteria.receivedBefore) parts.push(`before ${criteria.receivedBefore}`);
  if (criteria.readStatus && criteria.readStatus !== 'any') parts.push(criteria.readStatus);
  if (criteria.hasAttachments && criteria.hasAttachments !== 'any') parts.push(criteria.hasAttachments === 'yes' ? 'with attachments' : 'without attachments');
  if (criteria.importance && criteria.importance !== 'any') parts.push(`${criteria.importance} importance`);
  return parts.length > 0 ? parts.join(', ') : 'any message';
};

const fetchAllMatchingIds = async (accessToken, criteria) => {
  const filter = buildGraphFilter(criteria);
  const params = new URLSearchParams({
    $top: '50',
    $select: 'id',
  });
  if (filter) params.set('$filter', filter);

  let url = `https://graph.microsoft.com/v1.0/me/messages?${params.toString()}`;
  const ids = [];

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
      },
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph fetch failed: ${response.status} ${errorText}`);
    }
    const data = await response.json();
    const pageIds = Array.isArray(data.value) ? data.value.map((item) => item.id).filter(Boolean) : [];
    ids.push(...pageIds);
    url = data['@odata.nextLink'] || '';
  }

  return ids;
};

const createCountJob = (criteria, accessToken) => {
  const id = createJobId();
  const job = {
    id,
    kind: 'count',
    status: 'queued',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    totalIds: 0,
    deletedCount: 0,
    statusMessage: 'Queued',
    error: '',
    criteria,
    summary: summarizeCriteria(criteria),
    accessToken,
  };
  jobs.set(id, job);

  setTimeout(async () => {
    job.status = 'running';
    job.statusMessage = 'Counting matching messages...';
    job.updatedAt = new Date().toISOString();

    try {
      const ids = await fetchAllMatchingIds(accessToken, criteria);
      job.totalIds = ids.length;
      job.status = 'completed';
      job.statusMessage = `Found ${ids.length} matching messages.`;
      job.updatedAt = new Date().toISOString();
    } catch (jobError) {
      const message = jobError instanceof Error ? jobError.message : 'Unknown count error.';
      job.status = 'failed';
      job.error = message;
      job.statusMessage = `Failed: ${message}`;
    } finally {
      job.updatedAt = new Date().toISOString();
      job.accessToken = undefined;
    }
  }, 0);

  return job;
};

const deleteIdsInBatches = async (accessToken, ids, job) => {
  const batchSize = 25;
  const delayMs = 500;

  for (let index = 0; index < ids.length; index += batchSize) {
    const batch = ids.slice(index, index + batchSize);

    for (const id of batch) {
      let attempts = 0;
      while (attempts < 4) {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
          method: 'DELETE',
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        });

        if (response.status === 429 || response.status >= 500) {
          attempts += 1;
          const waitMs = 500 * attempts;
          await new Promise((resolve) => setTimeout(resolve, waitMs));
          continue;
        }

        if (!response.ok && response.status !== 202) {
          const text = await response.text();
          throw new Error(`Graph delete failed: ${response.status} ${text}`);
        }

        job.deletedCount += 1;
        job.statusMessage = `Deleted ${job.deletedCount}/${job.totalIds} messages...`;
        job.updatedAt = new Date().toISOString();
        break;
      }
    }

    if (index + batchSize < ids.length) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }
};

const createDeleteJob = (criteria, accessToken) => {
  const id = createJobId();
  const job = {
    id,
    kind: 'delete',
    status: 'queued',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    totalIds: 0,
    deletedCount: 0,
    statusMessage: 'Queued',
    error: '',
    criteria,
    summary: summarizeCriteria(criteria),
    accessToken,
  };
  jobs.set(id, job);

  setTimeout(async () => {
    job.status = 'running';
    job.statusMessage = 'Fetching matching messages...';
    job.updatedAt = new Date().toISOString();

    try {
      const ids = await fetchAllMatchingIds(accessToken, criteria);
      job.totalIds = ids.length;
      job.updatedAt = new Date().toISOString();

      if (ids.length === 0) {
        job.status = 'completed';
        job.statusMessage = 'No matching messages found.';
      } else {
        job.statusMessage = `Deleting ${ids.length} matching messages...`;
        await deleteIdsInBatches(accessToken, ids, job);
        job.status = 'completed';
        job.statusMessage = `Deleted ${ids.length} matching messages.`;
      }
    } catch (jobError) {
      const message = jobError instanceof Error ? jobError.message : 'Unknown deletion error.';
      job.status = 'failed';
      job.error = message;
      job.statusMessage = `Failed: ${message}`;
    } finally {
      job.updatedAt = new Date().toISOString();
      job.accessToken = undefined;
    }
  }, 0);

  return job;
};

const getSafeJobData = (job) => ({
  id: job.id,
  kind: job.kind || 'delete',
  status: job.status,
  createdAt: job.createdAt,
  updatedAt: job.updatedAt,
  totalIds: job.totalIds,
  deletedCount: job.deletedCount,
  statusMessage: job.statusMessage,
  error: job.error,
  summary: job.summary,
  criteria: job.criteria,
});

app.post('/jobs', (req, res) => {
  const { accessToken, criteria } = req.body;
  if (!accessToken || typeof accessToken !== 'string') {
    return res.status(400).json({ error: 'Missing accessToken in request body.' });
  }

  if (!criteria || typeof criteria !== 'object') {
    return res.status(400).json({ error: 'Missing deletion criteria.' });
  }

  const job = createDeleteJob(criteria, accessToken);
  return res.status(202).json({ id: job.id, status: job.status });
});

app.post('/jobs/count', (req, res) => {
  const { accessToken, criteria } = req.body;
  if (!accessToken || typeof accessToken !== 'string') {
    return res.status(400).json({ error: 'Missing accessToken in request body.' });
  }

  if (!criteria || typeof criteria !== 'object') {
    return res.status(400).json({ error: 'Missing count criteria.' });
  }

  const job = createCountJob(criteria, accessToken);
  return res.status(202).json({ id: job.id, status: job.status, totalIds: job.totalIds });
});

app.get('/jobs', (req, res) => {
  const jobList = Array.from(jobs.values()).map(getSafeJobData);
  return res.json(jobList);
});

app.get('/jobs/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) {
    return res.status(404).json({ error: 'Job not found.' });
  }
  return res.json(getSafeJobData(job));
});

// Reports endpoint and CSV parsing removed — mailbox-size feature reverted.

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

if (process.env.NODE_ENV !== 'test') {
  app.listen(port, () => {
    console.log(`API service listening on port ${port}`);
  });
}

export { app };
