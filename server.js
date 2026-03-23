const express = require('express');
const axios = require('axios');
const path = require('path');

// Load .env file if present
try { require('dotenv').config(); } catch (e) { /* dotenv optional */ }

const app = express();
const PORT = process.env.PORT || 3000;

// Zendesk configuration (from environment or defaults)
const ZENDESK_DOMAIN = process.env.ZENDESK_DOMAIN || 'b2c-innovation.zendesk.com';
const ZENDESK_EMAIL = process.env.ZENDESK_EMAIL || 'simone.carroccia@24hassistance.com/token';
const ZENDESK_TOKEN = process.env.ZENDESK_TOKEN || 'MSLsnRIbfU43NofdNryRFbVpNRo8zPghtX5SxTEE';
const ZENDESK_BASE = `https://${ZENDESK_DOMAIN}/api/v2`;

const zendesk = axios.create({
  baseURL: ZENDESK_BASE,
  auth: { username: ZENDESK_EMAIL, password: ZENDESK_TOKEN },
  headers: { 'Content-Type': 'application/json' },
  timeout: 30000
});

// Simple in-memory cache
const cache = {};
function getCached(key, ttlMs) {
  const entry = cache[key];
  if (entry && Date.now() - entry.ts < ttlMs) return entry.data;
  return null;
}
function setCache(key, data) {
  cache[key] = { data, ts: Date.now() };
}

// Retry on 429
async function zendeskGet(url, retries = 3) {
  for (let i = 0; i <= retries; i++) {
    try {
      return await zendesk.get(url);
    } catch (err) {
      if (err.response && err.response.status === 429 && i < retries) {
        const wait = (err.response.headers['retry-after'] || 10) * 1000;
        await new Promise(r => setTimeout(r, wait));
      } else {
        throw err;
      }
    }
  }
}

// Paginate through Zendesk search results
async function searchAll(query) {
  let results = [];
  let url = `/search.json?query=${encodeURIComponent(query)}&per_page=100`;
  while (url) {
    const res = await zendeskGet(url);
    results = results.concat(res.data.results || []);
    if (res.data.next_page) {
      // next_page is a full URL, convert to relative
      url = res.data.next_page.replace(ZENDESK_BASE, '');
    } else {
      url = null;
    }
    // Safety: Zendesk search max 1000 results
    if (results.length >= 1000) break;
  }
  return results;
}

// GET /api/agents
app.get('/api/agents', async (req, res) => {
  try {
    const cached = getCached('agents', 10 * 60 * 1000);
    if (cached) return res.json(cached);

    let agents = [];
    let url = '/users.json?role=agent&per_page=100';
    while (url) {
      const r = await zendeskGet(url);
      agents = agents.concat(r.data.users || []);
      url = r.data.next_page ? r.data.next_page.replace(ZENDESK_BASE, '') : null;
    }

    // Also fetch admins who may handle tickets
    let urlAdmin = '/users.json?role=admin&per_page=100';
    while (urlAdmin) {
      const r = await zendeskGet(urlAdmin);
      agents = agents.concat(r.data.users || []);
      urlAdmin = r.data.next_page ? r.data.next_page.replace(ZENDESK_BASE, '') : null;
    }

    const mapped = agents.map(u => ({
      id: u.id,
      name: u.name,
      email: u.email,
      photo: u.photo ? u.photo.content_url : null
    }));

    setCache('agents', mapped);
    res.json(mapped);
  } catch (err) {
    console.error('Error fetching agents:', err.message);
    res.status(500).json({ error: 'Errore nel recupero degli agenti' });
  }
});

// Fetch resolved tickets for a given date
async function getResolvedTickets(date) {
  const cacheKey = `resolved_${date}`;
  const cached = getCached(cacheKey, 2 * 60 * 1000);
  if (cached) return cached;

  const nextDay = new Date(date + 'T00:00:00Z');
  nextDay.setUTCDate(nextDay.getUTCDate() + 1);
  const nextDayStr = nextDay.toISOString().split('T')[0];

  const query = `type:ticket status:solved solved>=${date} solved<${nextDayStr}`;
  const tickets = await searchAll(query);

  // Also search for closed tickets solved today
  const queryClosed = `type:ticket status:closed solved>=${date} solved<${nextDayStr}`;
  const closedTickets = await searchAll(queryClosed);

  const allTickets = [...tickets, ...closedTickets];

  // Deduplicate by ticket id
  const seen = new Set();
  const unique = allTickets.filter(t => {
    if (seen.has(t.id)) return false;
    seen.add(t.id);
    return true;
  });

  // Try to fetch metrics in batch
  let metricsMap = {};
  if (unique.length > 0) {
    try {
      const ids = unique.map(t => t.id);
      // Batch in groups of 100
      for (let i = 0; i < ids.length; i += 100) {
        const batch = ids.slice(i, i + 100).join(',');
        const mRes = await zendeskGet(`/tickets/show_many.json?ids=${batch}&include=metric_sets`);
        if (mRes.data.metric_sets) {
          mRes.data.metric_sets.forEach(m => {
            metricsMap[m.ticket_id] = m;
          });
        }
      }
    } catch (e) {
      console.warn('Could not fetch ticket metrics:', e.message);
    }
  }

  const result = unique.map(t => {
    const metrics = metricsMap[t.id];
    const resolutionMinutes = metrics && metrics.full_resolution_time_in_minutes
      ? (metrics.full_resolution_time_in_minutes.business || metrics.full_resolution_time_in_minutes.calendar)
      : null;

    return {
      id: t.id,
      subject: t.subject,
      status: t.status,
      priority: t.priority,
      created_at: t.created_at,
      updated_at: t.updated_at,
      assignee_id: t.assignee_id,
      requester_id: t.requester_id,
      tags: t.tags || [],
      resolution_minutes: resolutionMinutes,
      url: `https://${ZENDESK_DOMAIN}/agent/tickets/${t.id}`
    };
  });

  setCache(cacheKey, result);
  return result;
}

// GET /api/tickets/resolved-today
app.get('/api/tickets/resolved-today', async (req, res) => {
  try {
    const date = req.query.date || new Date().toISOString().split('T')[0];
    const tickets = await getResolvedTickets(date);

    // Group by assignee
    const grouped = {};
    tickets.forEach(t => {
      const aid = t.assignee_id || 'unassigned';
      if (!grouped[aid]) grouped[aid] = [];
      grouped[aid].push(t);
    });

    res.json({ date, total: tickets.length, byAgent: grouped });
  } catch (err) {
    console.error('Error fetching resolved tickets:', err.message);
    res.status(500).json({ error: 'Errore nel recupero dei ticket risolti' });
  }
});

// GET /api/stats
app.get('/api/stats', async (req, res) => {
  try {
    const date = req.query.date || new Date().toISOString().split('T')[0];
    const tickets = await getResolvedTickets(date);

    // Fetch agents for name mapping
    const agentsCached = getCached('agents', 10 * 60 * 1000);
    let agentsMap = {};
    if (agentsCached) {
      agentsCached.forEach(a => { agentsMap[a.id] = a; });
    }

    // Per-agent stats
    const perAgent = {};
    tickets.forEach(t => {
      const aid = t.assignee_id || 'unassigned';
      if (!perAgent[aid]) perAgent[aid] = { count: 0, totalMinutes: 0, withMetrics: 0 };
      perAgent[aid].count++;
      if (t.resolution_minutes) {
        perAgent[aid].totalMinutes += t.resolution_minutes;
        perAgent[aid].withMetrics++;
      }
    });

    const perAgentArr = Object.entries(perAgent).map(([id, data]) => ({
      agentId: id,
      name: agentsMap[id] ? agentsMap[id].name : (id === 'unassigned' ? 'Non assegnato' : `Agente #${id}`),
      photo: agentsMap[id] ? agentsMap[id].photo : null,
      count: data.count,
      avgResolutionMinutes: data.withMetrics > 0 ? Math.round(data.totalMinutes / data.withMetrics) : null
    })).sort((a, b) => b.count - a.count);

    // Global stats
    const totalResolved = tickets.length;
    const topAgent = perAgentArr.length > 0 ? perAgentArr[0] : null;
    const allMinutes = tickets.filter(t => t.resolution_minutes).map(t => t.resolution_minutes);
    const avgResolution = allMinutes.length > 0
      ? Math.round(allMinutes.reduce((a, b) => a + b, 0) / allMinutes.length)
      : null;

    // Priority breakdown
    const byPriority = {};
    tickets.forEach(t => {
      const p = t.priority || 'none';
      byPriority[p] = (byPriority[p] || 0) + 1;
    });

    res.json({
      date,
      totalResolved,
      topAgent,
      avgResolutionMinutes: avgResolution,
      byPriority,
      perAgent: perAgentArr
    });
  } catch (err) {
    console.error('Error computing stats:', err.message);
    res.status(500).json({ error: 'Errore nel calcolo delle statistiche' });
  }
});

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => {
  console.log(`Dashboard avviata su http://localhost:${PORT}`);
});
