import { useMemo, useState, useEffect } from 'react';
import { createPortal } from 'react-dom';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider, useMsal, useIsAuthenticated } from '@azure/msal-react';

interface MessageItem {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments?: boolean;
  importance?: string;
}

interface JobStatus {
  id: string;
  kind?: 'count' | 'delete';
  status: 'queued' | 'running' | 'completed' | 'failed';
  createdAt: string;
  updatedAt: string;
  totalIds: number;
  deletedCount: number;
  statusMessage: string;
  error?: string;
  summary: string;
  criteria?: Record<string, any>;
}

const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID}`,
    redirectUri: import.meta.env.VITE_REDIRECT_URI || window.location.origin,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level: any, message: any, containsPii: any) => {
        if (containsPii) return;
        // eslint-disable-next-line no-console
        console.log(`[MSAL] ${message}`);
      },
    },
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

const loginRequest = {
  scopes: ['openid', 'profile', 'User.Read', 'Mail.ReadWrite'],
};

const graphScopeRequest = {
  scopes: ['User.Read', 'Mail.ReadWrite'],
};

const graphEndpoint = 'https://graph.microsoft.com/v1.0/me/messages';

function buildGraphFilter(
  sender: string,
  subject: string,
  body: string,
  folderId: string,
  receivedAfter: string,
  receivedBefore: string,
  readStatus: string,
  hasAttachments: string,
  importance: string,
  senderExclude: string = '',
  subjectExclude: string = '',
) {
  const filterParts: string[] = [];

  if (sender.trim()) {
    const safeSender = sender.trim().replace(/'/g, "''");
    filterParts.push(`contains(from/emailAddress/address,'${safeSender}')`);
  }

  if (senderExclude.trim()) {
    const safeSender = senderExclude.trim().replace(/'/g, "''");
    filterParts.push(`not(contains(from/emailAddress/address,'${safeSender}'))`);
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

  if (folderId.trim()) {
    const safeFolderId = folderId.trim().replace(/'/g, "''");
    filterParts.push(`parentFolderId eq '${safeFolderId}'`);
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
}

function prettyDate(value: string) {
  return new Date(value).toLocaleString();
}

function formatBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 KB';
  const units = ['B', 'KB', 'MB', 'GB'];
  let v = bytes;
  let i = 0;
  while (v >= 1024 && i < units.length - 1) { v /= 1024; i += 1; }
  return `${v >= 10 || i === 0 ? Math.round(v) : v.toFixed(1)} ${units[i]}`;
}


function stripHtml(value: string) {
  return value
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

function MainApp() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [sender, setSender] = useState('');
  const [subject, setSubject] = useState('');
  const [senderExclude, setSenderExclude] = useState('');
  const [subjectExclude, setSubjectExclude] = useState('');
  const [body, setBody] = useState('');
  const [folderId, setFolderId] = useState('');
  const [receivedAfter, setReceivedAfter] = useState('');
  const [receivedBefore, setReceivedBefore] = useState('');
  const [readStatus, setReadStatus] = useState('any');
  const [hasAttachments, setHasAttachments] = useState('any');
  const [importance, setImportance] = useState('any');
  const [messages, setMessages] = useState<MessageItem[]>([]);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [deleteStatus, setDeleteStatus] = useState<string | null>(null);
  const [jobs, setJobs] = useState<JobStatus[]>([]);
  const [nextPageLink, setNextPageLink] = useState<string | null>(null);
  const [pageHistory, setPageHistory] = useState<string[]>([]);
  const [currentPageUrl, setCurrentPageUrl] = useState<string | null>(null);
  const [messageBodies, setMessageBodies] = useState<Record<string, string>>({});
  const [expandedBodyId, setExpandedBodyId] = useState<string | null>(null);
  const [loadingBodyId, setLoadingBodyId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [mailboxStats, setMailboxStats] = useState<{ totalItemCount?: number; unreadItemCount?: number; folderCount?: number } | null>(null);
  const [folders, setFolders] = useState<{ id: string; displayName: string }[]>([]);
  const [totalMatchCount, setTotalMatchCount] = useState<number | null>(null);
  const [countStatus, setCountStatus] = useState<string | null>(null);
  

  const account = accounts[0] ?? null;

  const acquireToken = async () => {
    try {
      const result = await instance.acquireTokenSilent({
        ...graphScopeRequest,
        account: account || undefined,
      });
      return result.accessToken;
    } catch (silentError) {
      const result = await instance.acquireTokenPopup(graphScopeRequest);
      return result.accessToken;
    }
  };

  const handleLogin = async () => {
    try {
      await instance.loginRedirect(loginRequest);
      setError(null);
    } catch (loginError) {
      setError('Unable to sign in. Please try again.');
    }
  };

  const handleLogout = async () => {
    try {
      await instance.logoutPopup();
      setMessages([]);
      setSelectedIds([]);
      setError(null);
    } catch {
      setError('Unable to sign out.');
    }
  };

  const fetchMessages = async (direction: 'first' | 'next' | 'previous' = 'first') => {
    setError(null);
    setIsLoading(true);

    try {
      const token = await acquireToken();
      const buildFirstUrl = () => {
        const filter = buildGraphFilter(sender, subject, body, folderId, receivedAfter, receivedBefore, readStatus, hasAttachments, importance, senderExclude, subjectExclude);
        const params = new URLSearchParams({
          $top: '50',
          $select: 'subject,from,receivedDateTime,isRead,id,hasAttachments,importance',
        });
        if (filter) params.set('$filter', filter);
        return `${graphEndpoint}?${params.toString()}`;
      };

      let requestUrl = buildFirstUrl();

      if (direction === 'next' && nextPageLink) {
        setPageHistory((previous) => [...previous, currentPageUrl || buildFirstUrl()]);
        requestUrl = nextPageLink;
      } else if (direction === 'previous') {
        const previousUrl = pageHistory[pageHistory.length - 1];
        if (previousUrl) {
          setPageHistory((previous) => previous.slice(0, -1));
          requestUrl = previousUrl;
        } else {
          requestUrl = buildFirstUrl();
        }
      } else {
        setPageHistory([]);
        setCurrentPageUrl(null);
        setNextPageLink(null);
        requestUrl = buildFirstUrl();
      }

      const response = await fetch(requestUrl, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error(`Graph request failed: ${response.status}`);
      }
      const data = await response.json();
      setMessages(data.value ?? []);
      setSelectedIds([]);
      setCurrentPageUrl(requestUrl);
      setNextPageLink(data['@odata.nextLink'] || null);
    } catch (fetchError) {
      setError(fetchError instanceof Error ? fetchError.message : 'Failed to load messages.');
    } finally {
      setIsLoading(false);
    }
  };

  const fetchMailboxUsage = async () => {
    try {
      const token = await acquireToken();
      const response = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders?$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount&$top=200', {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error(`Mailbox stats request failed: ${response.status}`);
      }
      const data = await response.json();
      const folders = Array.isArray(data.value) ? data.value : [];
      const totalItemCount = folders.reduce((sum, folder) => sum + (Number(folder.totalItemCount) || 0), 0);
      const unreadItemCount = folders.reduce((sum, folder) => sum + (Number(folder.unreadItemCount) || 0), 0);
      setMailboxStats({
        totalItemCount,
        unreadItemCount,
        folderCount: folders.length,
      });
      setFolders(folders.map((folder: any) => ({ id: folder.id, displayName: folder.displayName || 'Unnamed' })));
    } catch (mailboxError) {
      // eslint-disable-next-line no-console
      console.error('Mailbox usage fetch failed:', mailboxError);
    }
  };


  const allSelected = messages.length > 0 && selectedIds.length === messages.length;

  const toggleSelectAll = () => {
    if (allSelected) {
      setSelectedIds([]);
      return;
    }
    setSelectedIds(messages.map((message) => message.id));
  };

  const deleteMessagesByIds = async (ids: string[]) => {
    const token = await acquireToken();
    const batchSize = 5;
    const delayMs = 400;

    for (let index = 0; index < ids.length; index += batchSize) {
      const batch = ids.slice(index, index + batchSize);

      for (const id of batch) {
        let attempts = 0;
        while (attempts < 4) {
          const response = await fetch(`${graphEndpoint}/${id}`, {
            method: 'DELETE',
            headers: {
              Authorization: `Bearer ${token}`,
            },
          });

          if (response.status === 429 || response.status >= 500) {
            attempts += 1;
            const waitMs = 500 * attempts;
            await new Promise((resolve) => setTimeout(resolve, waitMs));
            continue;
          }

          if (!response.ok && response.status !== 202) {
            throw new Error(`Graph delete failed: ${response.status}`);
          }
          break;
        }
      }

      if (index + batchSize < ids.length) {
        await new Promise((resolve) => setTimeout(resolve, delayMs));
      }
    }
  };

  const deleteSelected = async () => {
    if (selectedIds.length === 0) {
      setError('Select at least one message to delete.');
      return;
    }

    if (!window.confirm(`Delete ${selectedIds.length} selected message(s)? This cannot be undone.`)) {
      return;
    }

    setIsDeleting(true);
    setDeleteStatus('Deleting selected messages...');
    setError(null);

    try {
      await deleteMessagesByIds(selectedIds);
      await fetchMessages();
      setDeleteStatus('Selected messages deleted.');
    } catch (deleteError) {
      setError(deleteError instanceof Error ? deleteError.message : 'Delete request failed.');
    } finally {
      setIsDeleting(false);
      setTimeout(() => setDeleteStatus(null), 3000);
    }
  };

  const fetchAllMatchingIds = async () => {
    const token = await acquireToken();
    const filter = buildGraphFilter(sender, subject, body, folderId, receivedAfter, receivedBefore, readStatus, hasAttachments, importance, senderExclude, subjectExclude);
    const params = new URLSearchParams({
      $top: '50',
      $select: 'id',
    });
    if (filter) params.set('$filter', filter);

    let url = `${graphEndpoint}?${params.toString()}`;
    const allIds: string[] = [];

    while (url) {
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error(`Graph request failed: ${response.status}`);
      }
      const data = await response.json();
      const ids = Array.isArray(data.value) ? data.value.map((message: any) => message.id).filter(Boolean) : [];
      allIds.push(...ids);
      url = data['@odata.nextLink'] || '';
    }

    return allIds;
  };

  const fetchMessageBody = async (messageId: string) => {
    if (messageBodies[messageId]) {
      setExpandedBodyId((current) => (current === messageId ? null : messageId));
      return;
    }

    setLoadingBodyId(messageId);
    setError(null);

    try {
      const token = await acquireToken();
      const response = await fetch(`${graphEndpoint}/${messageId}?$select=body,subject`, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error(`Body fetch failed: ${response.status}`);
      }

      const data = await response.json();
      const bodyContent = data.body?.content || 'No message body returned.';
      setMessageBodies((previous) => ({ ...previous, [messageId]: bodyContent }));
      setExpandedBodyId(messageId);
    } catch (bodyError) {
      setError(bodyError instanceof Error ? bodyError.message : 'Failed to fetch message body.');
    } finally {
      setLoadingBodyId(null);
    }
  };

  const loadJobs = async () => {
    try {
      const response = await fetch(`${import.meta.env.VITE_API_BASE_URL}/jobs`);
      if (!response.ok) {
        throw new Error(`Job list request failed: ${response.status}`);
      }
      const data = await response.json();
      setJobs(Array.isArray(data) ? data : []);
    } catch (jobError) {
      // eslint-disable-next-line no-console
      console.error('Job list fetch failed:', jobError);
    }
  };

  const rerunJob = async (job: JobStatus) => {
    if (!job.criteria) {
      setError('Unable to rerun: original criteria not available.');
      return;
    }
    setDeleteStatus(`Re-running ${job.kind === 'count' ? 'count' : 'delete'} job...`);
    setError(null);
    try {
      const token = await acquireToken();
      const endpoint = job.kind === 'count' ? '/jobs/count' : '/jobs';
      const response = await fetch(`${import.meta.env.VITE_API_BASE_URL}${endpoint}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ accessToken: token, criteria: job.criteria }),
      });
      if (!response.ok) {
        const text = await response.text();
        throw new Error(`Rerun failed: ${response.status} ${text}`);
      }
      const data = await response.json();
      setDeleteStatus(`Re-run job created: ${data.id}`);
      await loadJobs();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to rerun job.');
    } finally {
      setTimeout(() => setDeleteStatus(null), 3000);
    }
  };

  const deleteAllMatching = async () => {
    if (!window.confirm('Create a background delete job for all matching messages?')) {
      return;
    }

    setIsDeleting(true);
    setDeleteStatus('Submitting background delete job...');
    setError(null);

    try {
      const token = await acquireToken();
      const selectedFolderName = folders.find((folder) => folder.id === folderId)?.displayName || '';
      const criteria = {
        sender,
        subject,
        senderExclude,
        subjectExclude,
        body,
        folder: folderId,
        folderName: selectedFolderName,
        receivedAfter,
        receivedBefore,
        readStatus,
        hasAttachments,
        importance,
      };

      const response = await fetch(`${import.meta.env.VITE_API_BASE_URL}/jobs`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ accessToken: token, criteria }),
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error(`Job create failed: ${response.status} ${text}`);
      }

      const data = await response.json();
      setDeleteStatus(`Background job created: ${data.id}`);
      await loadJobs();
    } catch (jobError) {
      setError(jobError instanceof Error ? jobError.message : 'Failed to create background delete job.');
    } finally {
      setIsDeleting(false);
      setTimeout(() => setDeleteStatus(null), 3000);
    }
  };

  const getTotalMessageCount = async () => {
    setError(null);
    setCountStatus('Starting total-message count job...');
    setTotalMatchCount(null);

    try {
      const token = await acquireToken();
      const selectedFolderName = folders.find((folder) => folder.id === folderId)?.displayName || '';
      const criteria = {
        sender,
        subject,
        senderExclude,
        subjectExclude,
        body,
        folder: folderId,
        folderName: selectedFolderName,
        receivedAfter,
        receivedBefore,
        readStatus,
        hasAttachments,
        importance,
      };

      const response = await fetch(`${import.meta.env.VITE_API_BASE_URL}/jobs/count`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ accessToken: token, criteria }),
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error(`Count job failed: ${response.status} ${text}`);
      }

      const data = await response.json();
      setCountStatus(`Counting messages in background... Job ${data.id}`);
      await loadJobs();

      const pollStatus = async () => {
        const jobResponse = await fetch(`${import.meta.env.VITE_API_BASE_URL}/jobs/${data.id}`);
        if (!jobResponse.ok) {
          throw new Error(`Unable to fetch count status: ${jobResponse.status}`);
        }

        const job = await jobResponse.json();
        if (job.status === 'completed') {
          setTotalMatchCount(job.totalIds ?? 0);
          setCountStatus(`Total matching messages: ${job.totalIds ?? 0}`);
          return;
        }

        if (job.status === 'failed') {
          throw new Error(job.error || 'Message count job failed.');
        }

        setCountStatus(`Counting... ${job.totalIds ?? 0} found so far`);
        setTimeout(pollStatus, 1500);
      };

      await pollStatus();
    } catch (countError) {
      setError(countError instanceof Error ? countError.message : 'Failed to count matching messages.');
      setCountStatus(null);
    }
  };

  useEffect(() => {
    if (!isAuthenticated) {
      setMailboxStats(null);
      setJobs([]);
      return;
    }
    fetchMailboxUsage();
    loadJobs();
  }, [isAuthenticated]);

  useEffect(() => {
    const activeJobExists = jobs.some((job) => job.status === 'queued' || job.status === 'running');
    if (!activeJobExists) return undefined;

    const interval = setInterval(() => {
      loadJobs();
    }, 5000);

    return () => clearInterval(interval);
  }, [jobs]);

  function MessageModal({ children, onClose }: { children: any; onClose: () => void }) {
    useEffect(() => {
      const originalOverflow = document.body.style.overflow;
      document.body.style.overflow = 'hidden';
      return () => {
        document.body.style.overflow = originalOverflow;
      };
    }, []);

    const modal = (
      <div className="modal-backdrop" onClick={onClose}>
        <div className="message-modal" onClick={(e) => e.stopPropagation()}>
          {children}
        </div>
      </div>
    );

    return createPortal(modal, document.body);
  }

  return (
    <div className="app-shell">
      <header>
        <div>
          <h1>Outlook Manager</h1>
          <p>Search, preview, and clean up Outlook mail with Microsoft Graph.</p>
        </div>
        <div className="auth-actions">
          {isAuthenticated ? (
            <button onClick={handleLogout}>Sign out</button>
          ) : (
            <button onClick={handleLogin}>Sign in with Microsoft</button>
          )}
        </div>
      </header>

      {!isAuthenticated ? (
        <section className="box">
          <h2>Sign in to get started</h2>
          <p>Use your Microsoft account to access Outlook cleanup tools.</p>
        </section>
      ) : (
        <>
          {mailboxStats ? (
            <section className="box mailbox-box">
              <h2>Mailbox summary</h2>
              <div className="mailbox-grid">
                <div className="stat-pill">
                  <span>Total messages</span>
                  <strong>{mailboxStats.totalItemCount ?? 0}</strong>
                </div>
                <div className="stat-pill">
                  <span>Unread messages</span>
                  <strong>{mailboxStats.unreadItemCount ?? 0}</strong>
                </div>
                <div className="stat-pill">
                  <span>Folders</span>
                  <strong>{mailboxStats.folderCount ?? 0}</strong>
                </div>
              </div>
            </section>
          ) : null}

          <section className="box">
            <h2>Advanced message filters</h2>
            <div className="filter-grid">
              <label>
                Sender contains
                <input value={sender} onChange={(event) => setSender(event.target.value)} placeholder="example@domain.com" />
              </label>
              <label>
                Sender does not contain
                <input value={senderExclude} onChange={(event) => setSenderExclude(event.target.value)} placeholder="exclude@domain.com" />
              </label>
              <label>
                Subject contains
                <input value={subject} onChange={(event) => setSubject(event.target.value)} placeholder="invoice, newsletter" />
              </label>
              <label>
                Subject does not contain
                <input value={subjectExclude} onChange={(event) => setSubjectExclude(event.target.value)} placeholder="promo, upsell" />
              </label>
              <label>
                Body contains
                <input value={body} onChange={(event) => setBody(event.target.value)} placeholder="receipt, meeting, promo" />
              </label>
              <label>
                Folder
                <select value={folderId} onChange={(event) => setFolderId(event.target.value)}>
                  <option value="">Any folder</option>
                  {folders.map((folder) => (
                    <option key={folder.id} value={folder.id}>{folder.displayName}</option>
                  ))}
                </select>
              </label>
              <label>
                Received after
                <input type="date" value={receivedAfter} onChange={(event) => setReceivedAfter(event.target.value)} />
              </label>
              <label>
                Received before
                <input type="date" value={receivedBefore} onChange={(event) => setReceivedBefore(event.target.value)} />
              </label>
              <label>
                Read status
                <select value={readStatus} onChange={(event) => setReadStatus(event.target.value)}>
                  <option value="any">Any</option>
                  <option value="read">Read</option>
                  <option value="unread">Unread</option>
                </select>
              </label>
              <label>
                Attachments
                <select value={hasAttachments} onChange={(event) => setHasAttachments(event.target.value)}>
                  <option value="any">Any</option>
                  <option value="yes">With attachments</option>
                  <option value="no">Without attachments</option>
                </select>
              </label>
              <label>
                Importance
                <select value={importance} onChange={(event) => setImportance(event.target.value)}>
                  <option value="any">Any</option>
                  <option value="low">Low</option>
                  <option value="normal">Normal</option>
                  <option value="high">High</option>
                </select>
              </label>
            </div>
            <div className="button-row">
              <button onClick={fetchMessages} disabled={isLoading}><span aria-hidden="true">🔎</span> Preview matching messages</button>
              <button onClick={getTotalMessageCount} disabled={isLoading || isDeleting}><span aria-hidden="true">📊</span> Get total message count</button>
            </div>
            {countStatus ? <div className="status-message">{countStatus}</div> : null}
            {totalMatchCount !== null ? (
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 8 }}>
                <div className="count-badge">Matching messages: <strong>{totalMatchCount}</strong></div>
                <button onClick={deleteAllMatching} disabled={isDeleting || isLoading} className="icon-button danger" title="Delete all matching (background)">
                  <span style={{ fontSize: 14 }}>🧹</span>
                </button>
              </div>
            ) : null}
          </section>

          <section className="box">
            <h2>Message preview</h2>
            {error ? <div className="error-message">{error}</div> : null}
            {deleteStatus ? <div className="status-message">{deleteStatus}</div> : null}
            {isLoading && !deleteStatus ? <div className="status-message">Loading messages...</div> : null}
            {messages.length === 0 ? <p>No messages loaded yet. Use preview to fetch results.</p> : (
              <>
                    <div className="button-row preview-tools">
                  <button type="button" onClick={toggleSelectAll}>
                    {allSelected ? 'Clear selection' : 'Select all'}
                  </button>
                  <button type="button" onClick={() => fetchMessages('previous')} disabled={isDeleting || isLoading || pageHistory.length === 0}>
                    Previous 50
                  </button>
                  <button type="button" onClick={() => fetchMessages('next')} disabled={isDeleting || isLoading || !nextPageLink}>
                    Next 50
                  </button>
                  <button type="button" onClick={deleteSelected} disabled={isDeleting || isLoading || selectedIds.length === 0}>
                    Delete selected
                  </button>
                  <button type="button" onClick={deleteAllMatching} disabled={isDeleting || isLoading}>
                    🧹 Delete all matching
                  </button>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Select</th>
                        <th>Subject</th>
                        <th>Sender</th>
                        <th>Received</th>
                        <th>Attachments</th>
                        <th>Importance</th>
                        <th>Read</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {messages.map((message) => (
                        <tr key={message.id}>
                          <td>
                            <input
                              type="checkbox"
                              checked={selectedIds.includes(message.id)}
                              onChange={(event) => {
                                const next = event.target.checked
                                  ? [...selectedIds, message.id]
                                  : selectedIds.filter((id) => id !== message.id);
                                setSelectedIds(next);
                              }}
                            />
                          </td>
                          <td>
                            <div>{message.subject || '(No subject)'}</div>
                          </td>
                          <td>{message.from?.emailAddress?.address || 'Unknown'}</td>
                          <td>{prettyDate(message.receivedDateTime)}</td>
                          <td>{message.hasAttachments ? 'Yes' : 'No'}</td>
                          <td>{message.importance || 'Normal'}</td>
                          <td>{message.isRead ? 'Yes' : 'No'}</td>
                          <td>
                            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                              <button
                                type="button"
                                className="icon-button"
                                onClick={() => fetchMessageBody(message.id)}
                                disabled={loadingBodyId === message.id}
                                aria-label={expandedBodyId === message.id ? 'Hide message body' : 'View message body'}
                                title={expandedBodyId === message.id ? 'Hide body' : 'View body'}
                              >
                                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden>
                                  <path d="M12 5c-7 0-11 7-11 7s4 7 11 7 11-7 11-7-4-7-11-7zm0 11a4 4 0 1 1 0-8 4 4 0 0 1 0 8z" fill="currentColor" />
                                </svg>
                              </button>

                              <button
                                type="button"
                                className="icon-button danger"
                                onClick={async () => {
                                  if (!window.confirm('Delete this message now? This cannot be undone.')) return;
                                  try {
                                    setIsDeleting(true);
                                    setDeleteStatus('Deleting message...');
                                    await deleteMessagesByIds([message.id]);
                                    await fetchMessages();
                                    setDeleteStatus('Message deleted.');
                                  } catch (err) {
                                    setError(err instanceof Error ? err.message : 'Failed to delete message');
                                  } finally {
                                    setIsDeleting(false);
                                    setTimeout(() => setDeleteStatus(null), 3000);
                                  }
                                }}
                                aria-label="Delete this message"
                                title="Delete this message"
                              >
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden>
                                  <path d="M3 6h18M8 6v12a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2V6" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
                                  <path d="M10 11v6M14 11v6M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
                                </svg>
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {expandedBodyId ? (
                  <MessageModal onClose={() => setExpandedBodyId(null)}>
                    <div className="message-modal-header">
                      <h3>Message body</h3>
                      <button type="button" onClick={() => setExpandedBodyId(null)}>Close</button>
                    </div>
                    <div className="message-body-text">
                      {stripHtml(messageBodies[expandedBodyId] || 'No message body available.')}
                    </div>
                  </MessageModal>
                ) : null}
              </>
            )}
          </section>

          <section className="box">
            <h2>Background delete jobs</h2>
            <p>These jobs run on the server and can continue if you close the browser.</p>
            <div className="button-row">
              <button type="button" onClick={loadJobs} disabled={isDeleting || isLoading}>
                Refresh job status
              </button>
            </div>
            {jobs.length === 0 ? (
              <p>No background delete jobs yet.</p>
            ) : (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Job</th>
                      <th>Status</th>
                      <th>Progress</th>
                      <th>Summary</th>
                      <th>Updated</th>
                      <th style={{ width: 60 }}>Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {jobs.map((job) => (
                      <tr key={job.id}>
                        <td>{job.kind === 'count' ? 'Count job' : 'Delete job'} · {job.id}</td>
                        <td>{job.status}</td>
                        <td>{job.deletedCount}/{job.totalIds}</td>
                        <td>{job.summary}</td>
                        <td>{new Date(job.updatedAt).toLocaleString()}</td>
                        <td>
                          {job.criteria ? (
                            <button
                              type="button"
                              className="icon-button"
                              onClick={() => rerunJob(job)}
                              aria-label="Rerun job with same criteria"
                              title="Rerun job with same criteria"
                            >
                              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden>
                                <path d="M21 12a9 9 0 1 1-3-6.7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                <path d="M21 3v6h-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              </svg>
                            </button>
                          ) : null}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </>
      )}
    </div>
  );
}

export default function App() {
  const pca = useMemo(() => new PublicClientApplication(msalConfig), []);

  useEffect(() => {
    (async () => {
      try {
        const result = await pca.handleRedirectPromise();
        // eslint-disable-next-line no-console
        console.log('Top-level MSAL handleRedirectPromise result:', result);
      } catch (e) {
        // eslint-disable-next-line no-console
        console.error('Top-level MSAL handleRedirectPromise error:', e);
      }
    })();
  }, [pca]);

  return (
    <MsalProvider instance={pca}>
      <MainApp />
    </MsalProvider>
  );
}
