/**
 * LDP LOGISTICS SECURITY DASHBOARD - UTILITIES
 */

// 1. Live Ticking Clock Utility
export const getLiveTime = () => {
  const now = new Date();
  return {
    time: now.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true
    }),
    date: now.toLocaleDateString('en-US', {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    })
  };
};

// 2. Severity classification
// Audit operations triaged the way a SOC would: destruction first, then exposure,
// then failed access, then routine work. The last rule matches everything, so an
// operation we have never seen is reported as Informational rather than dropped
// from the totals - a breakdown that silently discards rows is worse than none.
//
// Two of these were caught against the real feed: "Recycled" is a deletion under
// another name, and "LoggedIn" never contains the substring "login". Missing them
// left Critical permanently empty and buried 329 sign-ins in Informational.
export const SEVERITY_RULES = [
  { key: 'critical', label: 'Critical', color: 'var(--sev-critical)', test: /delete|recycl|malware|ransom|purge/i },
  { key: 'high', label: 'High', color: 'var(--sev-high)', test: /sharing|anonymouslink|companylink|addedtogroup|permission|siteadmin|roleassign/i },
  { key: 'medium', label: 'Medium', color: 'var(--sev-medium)', test: /failed|download|moved|renamed|restore/i },
  { key: 'low', label: 'Low', color: 'var(--sev-low)', test: /accessed|viewed|modified|uploaded|checkin|checkout|logged|login|signin|preview/i },
  { key: 'info', label: 'Informational', color: 'var(--sev-info)', test: /.*/ },
];

export const severityOf = (operation = '') =>
  (SEVERITY_RULES.find(rule => rule.test.test(operation)) || SEVERITY_RULES[SEVERITY_RULES.length - 1]).key;

// Service principals are not people and only add noise to a feed meant to be
// scanned by a human.
const isServiceAccount = (user = '') =>
  /^serviceprincipal_|app@sharepoint|urn:spo|^sharepoint\\/i.test(user);

const FEED_STATUS = { critical: 'alert', high: 'warn', medium: 'warn', low: 'ok', info: 'ok' };

// 3. Activity Feed
// Returns the most RECENT entries. This used to slice from the front of the
// array, which is the oldest end of the window - a "live" feed showing the
// oldest events it knows about. It also used to fall back to invented sample
// rows when there was no data, so an empty feed looked like a busy one.
export const generateActivityFeed = (rawData = [], limit = 20) =>
  rawData
    .filter(item => !isServiceAccount(item.UserId || ''))
    .slice()
    .sort((a, b) => new Date(b.CreationTime) - new Date(a.CreationTime))
    .slice(0, limit)
    .map((item, index) => ({
      id: item.Id || `${item.CreationTime}-${index}`,
      user: item.UserId?.split('@')[0] || 'System',
      action: item.Operation || 'Access',
      timestamp: new Date(item.CreationTime).toLocaleTimeString(),
      status: FEED_STATUS[severityOf(item.Operation)] || 'ok',
    }));
