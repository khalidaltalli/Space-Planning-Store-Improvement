/* ════════════════════════════════════════════════════════════
   WORK HUB — shared API helper.
   Include this on every workhub/*.html page before your page script.
   ════════════════════════════════════════════════════════════ */

// PASTE your Work Hub Apps Script Web App URL here (the /exec URL
// from deploying WorkHub_Code.gs as its own deployment).
const WORKHUB_API_URL = 'https://script.google.com/macros/s/AKfycbxW3TlLy7D90pN226hz7NExxv14gRjlWP418r2uSGTG0PZcQedKguI9CvtjGon3-Ntw/exec';

const WorkHub = (function () {
  // Token comes from the URL fragment the first time, e.g.
  // https://you.github.io/workhub/#token=xxxxxxxx
  // then it's cached in sessionStorage for the rest of the visit.
  function getToken() {
    let token = sessionStorage.getItem('wh_token');
    if (!token && location.hash.startsWith('#token=')) {
      token = decodeURIComponent(location.hash.replace('#token=', ''));
      sessionStorage.setItem('wh_token', token);
      // strip token out of the visible URL bar
      history.replaceState(null, '', location.pathname + location.search);
    }
    return token || '';
  }

  async function get(action, extraParams) {
    const token = getToken();
    const params = new URLSearchParams({ action, token, ...(extraParams || {}) });
    const res = await fetch(`${WORKHUB_API_URL}?${params.toString()}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    return data;
  }

  async function post(body) {
    const token = getToken();
    const res = await fetch(WORKHUB_API_URL, {
      method: 'POST',
      // text/plain avoids a CORS preflight that Apps Script can't handle
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...body, token })
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    return data;
  }

  async function requireAccess() {
    if (!getToken()) {
      document.body.innerHTML = '<p style="font-family:sans-serif;padding:40px;">Access link missing or invalid.</p>';
      throw new Error('no token');
    }
    try {
      await get('getBranches');
    } catch (err) {
      document.body.innerHTML = '<p style="font-family:sans-serif;padding:40px;">Access denied.</p>';
      throw err;
    }
  }

  function getBranches() {
    return get('getBranches').then(d => d.rows || []);
  }

  function getTab(tab, branch) {
    return get('getTab', branch ? { tab, branch } : { tab }).then(d => d.rows || []);
  }

  function saveRow(tab, fields, id) {
    return post({ action: 'saveRow', tab, fields, id: id || '' });
  }

  function deleteRow(tab, id) {
    return post({ action: 'deleteRow', tab, id });
  }

  return { requireAccess, getBranches, getTab, saveRow, deleteRow };
})();
