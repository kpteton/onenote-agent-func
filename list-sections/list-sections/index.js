module.exports = async function (context, req) {
  try {
    // require inside the handler so missing modules surface as a clear 500
    const { ConfidentialClientApplication } = require("@azure/msal-node");

    // Node 18+ has global fetch
    async function gfetch(url, token, init = {}) {
      const r = await fetch(url, {
        ...init,
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(init.headers || {}) }
      });
      if (r.status === 429 || r.status === 503) {
        const ra = r.headers.get("Retry-After");
        const ms = ra ? Number(ra) * 1000 : 2000;
        await new Promise(d => setTimeout(d, ms));
        return gfetch(url, token, init);
      }
      if (!r.ok) throw new Error(`Graph ${r.status} ${r.statusText}: ${await r.text()}`);
      const ct = r.headers.get("content-type") || "";
      return ct && ct.includes("application/json") ? r.json() : r.text();
    }

    const TENANT_ID     = process.env.TENANT_ID;
    const CLIENT_ID     = process.env.CLIENT_ID;
    const CLIENT_SECRET = process.env.CLIENT_SECRET;
    if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
      context.res = { status: 500, body: { error: "Server configuration missing required environment variables." } };
      return;
    }

    // Expect a user bearer token (Copilot will provide this later)
    const auth = (req.headers.authorization || req.headers.Authorization || "");
    const userJwt = auth.replace(/^Bearer\s+/i, "");
    if (!userJwt) { context.res = { status: 401, body: { error: "Missing user bearer token" } }; return; }

    const msal = new ConfidentialClientApplication({
      auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, clientSecret: CLIENT_SECRET }
    });
    const obo = await msal.acquireTokenOnBehalfOf({
      oboAssertion: userJwt,
      scopes: ["https://graph.microsoft.com/.default"]
    });
    if (!obo?.accessToken) throw new Error("Failed to get Graph token via OBO");
    const token = obo.accessToken;

    const { siteUrl, notebookName } = req.body || {};
    if (!notebookName) {
      context.res = { status: 400, body: { error: "notebookName required" } };
      return;
    }

    // Resolve site if provided
    let siteId = null;
    if (siteUrl) {
      const path = new URL(siteUrl).pathname; // e.g., /sites/TetonSales
      const site = await gfetch(`https://graph.microsoft.com/v1.0/sites/root:${encodeURI(path)}`, token);
      siteId = site.id;
    }

    // Find notebook
    const nbUrl = siteId
      ? `https://graph.microsoft.com/v1.0/sites/${siteId}/onenote/notebooks?$top=200`
      : `https://graph.microsoft.com/v1.0/me/onenote/notebooks?$top=200`;

    const nbs = await gfetch(nbUrl, token);
    const nb = (nbs.value || []).find(n => (n.displayName || "").trim().toLowerCase() === notebookName.trim().toLowerCase());
    if (!nb) { context.res = { status: 404, body: { error: "Notebook not found" } }; return; }

    // Sections
    const secUrl = siteId
      ? `https://graph.microsoft.com/v1.0/sites/${siteId}/onenote/notebooks/${nb.id}/sections?$top=200`
      : `https://graph.microsoft.com/v1.0/me/onenote/notebooks/${nb.id}/sections?$top=200`;

    const secs = await gfetch(secUrl, token);
    const sections = (secs.value || []).map(s => ({
      sectionId: s.id,
      sectionName: s.displayName,
      notebookId: nb.id,
      notebookName: nb.displayName
    }));
    context.res = { status: 200, body: sections };
  } catch (e) {
    context.log("list-sections error:", e?.message || e);
    context.res = { status: 500, body: { error: e?.message || "Unknown error" } };
  }
};
