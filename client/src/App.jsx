// v1.0.3 - Restored clean structure after sync errors
import React, { useState, useEffect, useMemo } from 'react';
import {
  Activity, Users, FileText, Share2, LayoutDashboard, Database,
  ShieldAlert, Laptop, Lock, Settings,
  ShieldCheck, ShieldX, Search, Trash2, AlertTriangle, RefreshCw, Globe, Power
} from 'lucide-react';
import {
  Chart as ChartJS, Tooltip as ChartTooltip, Legend as ChartLegend,
  CategoryScale, LinearScale, BarElement, Title, PointElement, LineElement, Filler
} from 'chart.js';
import { Bar as BarChartJS, Line as LineChartJS } from 'react-chartjs-2';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import { generateActivityFeed, SEVERITY_RULES, severityOf } from './dashboard';
import './App.css';

ChartJS.register(ChartTooltip, ChartLegend, CategoryScale, LinearScale, BarElement, Title, PointElement, LineElement, Filler);

// The backend polls Microsoft 365 every 15s, so refreshing faster than that just
// re-fetches an unchanged payload.
const REFRESH_MS = 10000;

// One list drives both the icon rail and the pill tabs, so the two can never
// drift out of sync.
const TABS = [
  { key: 'dashboard', label: 'Overview', icon: LayoutDashboard },
  { key: 'users', label: 'Users', icon: Users },
  { key: 'behavior', label: 'Security Pulse', icon: ShieldAlert },
  { key: 'devices', label: 'Devices', icon: Laptop },
  { key: 'leak', label: 'Data Leak', icon: ShieldX },
  { key: 'web', label: 'Web Usage', icon: Globe },
  { key: 'admin', label: 'Admin', icon: Settings },
];

const formatWhen = (value) => {
  const at = new Date(value);
  return Number.isNaN(at.getTime()) ? '—' : at.toLocaleString();
};

const LiveActivityFeed = ({ data }) => {
    const feed = useMemo(() => generateActivityFeed(data), [data]);
    return (
        <div className="live-feed-panel">
            <div className="feed-header">
                <h3 className="chart-title">
                    <span className="pulse-dot cyan"></span>
                    Live Activity Feed
                </h3>
                <span className="nav-badge" style={{ background: 'var(--accent-primary)', color: '#000' }}>LIVE</span>
            </div>
            <div className="feed-scroll">
                {feed.map(item => (
                    <div key={item.id} className="feed-item">
                        <span className={`pulse-dot ${item.status === 'alert' ? 'red' : item.status === 'warn' ? 'yellow' : 'green'}`}></span>
                        <div className="feed-info">
                            <div className="feed-user-action">
                                <span style={{ color: 'var(--accent-primary)' }}>{item.user}</span>
                                <span style={{ color: 'var(--text-muted)', margin: '0 8px', fontWeight: '400' }}>•</span>
                                <span>{item.action}</span>
                            </div>
                            <span className="feed-timestamp">{item.timestamp}</span>
                        </div>
                        <span className={`feed-badge ${item.status}`} style={{ marginLeft: 'auto' }}>
                            {item.status.toUpperCase()}
                        </span>
                    </div>
                ))}
            </div>
        </div>
    );
};

// Every panel polls on its own schedule, but an operator who has just changed
// something wants to see the result now rather than wait out the interval.
const RefreshButton = ({ onRefresh, busy }) => (
  <button className="refresh-btn" onClick={onRefresh} disabled={busy} title="Refresh this panel now">
    <RefreshCw size={13} className={busy ? 'spinning' : ''} />
    {busy ? 'Syncing' : 'Refresh'}
  </button>
);

// Panel title on the left, refresh on the right.
const PanelHead = ({ icon: Icon, title, onRefresh, busy }) => (
  <div className="panel-head">
    <div className="chart-title" style={{ marginBottom: 0 }}>
      {Icon && <Icon size={18} color="var(--accent-primary)" />} {title}
    </div>
    {onRefresh && <RefreshButton onRefresh={onRefresh} busy={busy} />}
  </div>
);

// The tag on a tile states the same proportion its blocks draw, so the meter is
// never decorative - both come from the two real counts.
const shareOf = (part, whole) => (whole > 0 ? `${Math.round((part / whole) * 100)}%` : null);

// Blocks rather than a smooth bar: at this size a reader can count them, which
// is the whole point of showing a proportion.
const Segments = ({ used, total, blocks = 8 }) => {
  const filled = total > 0 ? Math.min(blocks, Math.round((used / total) * blocks)) : 0;
  return (
    <div className="segments" aria-hidden="true">
      {Array.from({ length: blocks }, (_, i) => (
        <span key={i} className={`segment ${i < filled ? '' : 'empty'}`} />
      ))}
    </div>
  );
};

// Pass total to get the segmented meter; leave it out for a bare figure.
const Tile = ({ icon: Icon, label, tag, value, unit, note, tone = '', used, total }) => (
  <div className={`tile ${tone} fadeIn`}>
    <div className="tile-top">
      <span className="tile-label">
        <span className="tile-icon">{Icon && <Icon size={15} />}</span>
        {label}
      </span>
      {tag && <span className="tile-tag">{tag}</span>}
    </div>
    <div className="tile-figure">
      {value}{unit !== undefined && <span> / {unit}</span>}
    </div>
    {note && <div className="tile-note">{note}</div>}
    {total !== undefined && <Segments used={used} total={total} />}
  </div>
);

// Severity as labelled rows, not a pie. Critical/High/Medium sit on neighbouring
// red-orange-amber hues that no reader separates reliably, so the name and the
// count carry the meaning and the bar only shows relative weight.
const SeverityBreakdown = ({ breakdown, total }) => {
  if (!total) return <div className="data-empty">No audit events in the current window.</div>;
  const peak = Math.max(...breakdown.map(s => s.count), 1);
  return (
    <div className="sev-list">
      {breakdown.map(s => (
        <div className="sev-row" key={s.key}>
          <span className="sev-name">
            <span className="sev-swatch" style={{ background: s.color }} />
            {s.label}
          </span>
          <div className="sev-track">
            <div className="sev-fill" style={{ width: `${(s.count / peak) * 100}%`, background: s.color }} />
          </div>
          <span className="sev-count">{s.count}</span>
        </div>
      ))}
    </div>
  );
};

const OperationsBreakdown = ({ operations }) => {
  if (!operations.length) return <div className="data-empty">No operations recorded yet.</div>;
  const peak = Math.max(...operations.map(o => o.count), 1);
  return (
    <div className="sev-list">
      {operations.map(op => (
        <div className="sev-row" key={op.name}>
          <span className="sev-name" title={op.name}>
            <span className="sev-swatch" style={{ background: op.color }} />
            {op.name}
          </span>
          <div className="sev-track">
            <div className="sev-fill" style={{ width: `${(op.count / peak) * 100}%`, background: op.color }} />
          </div>
          <span className="sev-count">{op.count}</span>
        </div>
      ))}
    </div>
  );
};

const DashboardContent = React.memo(({ data, metrics, severityBreakdown, topOperations, activityTimeline, entitiesData, chartColors, onRefresh, busy }) => {
  return (
    <div className="fadeIn">
      <div className="panel-head">
        <p className="panel-caption" style={{ margin: 0 }}>Most recent {metrics.total.toLocaleString()} audit events. Security Pulse scores the full 24-hour retention window, so its counts are larger.</p>
        <RefreshButton onRefresh={onRefresh} busy={busy} />
      </div>
      <div className="tile-row">
        <Tile icon={Database} label="Events" value={metrics.total.toLocaleString()} note={`Across ${metrics.users} identities`} />
        <Tile tone="blush" icon={AlertTriangle} label="Destructive" value={metrics.deletions.toLocaleString()} tag={shareOf(metrics.deletions, metrics.total)} note="Deletions and purges" used={metrics.deletions} total={metrics.total} />
        <Tile tone="sky" icon={Share2} label="Exposure" value={metrics.shares.toLocaleString()} tag={shareOf(metrics.shares, metrics.total)} note="Sharing and link grants" used={metrics.shares} total={metrics.total} />
        <Tile tone="mint" icon={Lock} label="Failed sign-ins" value={metrics.failedLogins.toLocaleString()} tag={shareOf(metrics.failedLogins, metrics.total)} note="Rejected authentication" used={metrics.failedLogins} total={metrics.total} />
      </div>

      <div className="grid-analysis">
        <div className="chart-panel">
          <div className="chart-title"><ShieldAlert size={18} color="var(--accent-primary)" /> Events by Severity</div>
          <SeverityBreakdown breakdown={severityBreakdown} total={metrics.total} />
        </div>

        <div className="chart-panel">
          <div className="chart-title"><Activity size={18} color="var(--accent-primary)" /> Event Volume, Last 12 Hours</div>
          <div style={{ height: '240px' }}>
            {activityTimeline.length ? (
              <LineChartJS
                data={{
                  labels: activityTimeline.map(p => p.hour),
                  datasets: [{
                    label: 'Events',
                    data: activityTimeline.map(p => p.count),
                    borderColor: chartColors.accent,
                    backgroundColor: chartColors.accentFill,
                    borderWidth: 2,
                    fill: true,
                    tension: 0.35,
                    pointRadius: 0,
                    pointHoverRadius: 5,
                    pointHoverBackgroundColor: chartColors.accent,
                  }]
                }}
                options={{
                  maintainAspectRatio: false,
                  interaction: { mode: 'index', intersect: false },
                  plugins: { legend: { display: false } },
                  scales: {
                    x: { grid: { display: false }, border: { color: chartColors.grid }, ticks: { color: chartColors.text, font: { size: 11 }, maxRotation: 0, autoSkipPadding: 16 } },
                    y: { beginAtZero: true, grid: { color: chartColors.grid }, border: { display: false }, ticks: { color: chartColors.text, font: { size: 11 }, precision: 0 } }
                  }
                }}
              />
            ) : <div className="data-empty">Waiting for events.</div>}
          </div>
        </div>

        <div className="chart-panel">
          <div className="chart-title"><FileText size={18} color="var(--accent-primary)" /> Top Operations</div>
          <OperationsBreakdown operations={topOperations} />
        </div>
      </div>

      <div className="grid-bottom">
        <div className="chart-panel">
          <div className="chart-title"><Users size={18} color="var(--accent-primary)" /> Most Active Identities</div>
          <div style={{ height: '320px' }}>
            {entitiesData.length ? (
              <BarChartJS
                data={{
                  labels: entitiesData.map(e => e.name),
                  datasets: [{
                    label: 'Events',
                    data: entitiesData.map(e => e.count),
                    backgroundColor: chartColors.bar,
                    borderRadius: 10, barThickness: 22, hoverBackgroundColor: chartColors.accent
                  }]
                }}
                options={{
                  indexAxis: 'y', maintainAspectRatio: false,
                  plugins: { legend: { display: false } },
                  scales: {
                    x: { grid: { color: chartColors.grid }, border: { display: false }, ticks: { color: chartColors.text, precision: 0 } },
                    y: { grid: { display: false }, border: { display: false }, ticks: { color: chartColors.text, font: { size: 11 }, padding: 8 } }
                  }
                }}
              />
            ) : <div className="data-empty">No identity activity yet.</div>}
          </div>
        </div>
        <div className="chart-panel">
          <LiveActivityFeed data={data} />
        </div>
      </div>
    </div>
  );
});

const RISK_PILL = { Critical: 'pill-red', Moderate: 'pill-yellow', Low: 'pill-green' };

const SecurityPulse = ({ riskStats, onRefresh, busy }) => (
  <div className="fadeIn">
    <div className="chart-panel">
      <PanelHead icon={ShieldAlert} title="User Risk Scoring" onRefresh={onRefresh} busy={busy} />
      <p className="panel-caption">
        Scored across every event retained in the last 24 hours. Each signal is capped, so one noisy
        category cannot by itself push an account to Critical.
      </p>
      {riskStats.length === 0 ? (
        <div className="data-empty">
          No scored identities yet.<br />
          Risk is calculated from deletions and sharing activity in the audit feed.
        </div>
      ) : (
        <div className="table-scroll">
          <table className="custom-table">
            <thead>
              <tr>
                <th>Identity</th>
                <th>Risk</th>
                <th>Score</th>
                <th>Events</th>
                <th>Files touched</th>
                <th>Indicators</th>
              </tr>
            </thead>
            <tbody>
              {riskStats.map(entry => (
                <tr key={entry.user}>
                  <td>
                    <div style={{ fontWeight: 700, color: 'var(--text-primary)' }}>{entry.user?.split('@')[0]}</div>
                    <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{entry.user}</div>
                  </td>
                  <td><span className={`pill ${RISK_PILL[entry.level] || 'pill-gray'}`}>{entry.level}</span></td>
                  <td className="num" style={{ fontWeight: 700 }}>{entry.score}</td>
                  <td className="num">{entry.activityCount}</td>
                  <td className="num">{entry.fileCount}</td>
                  <td style={{ color: 'var(--text-secondary)', fontSize: '0.75rem' }}>
                    {entry.flags?.length ? entry.flags.join(' · ') : '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  </div>
);

const LEAK_PILL = { critical: 'pill-red', high: 'pill-orange', medium: 'pill-yellow' };
// Mirrors the thresholds in the backend's detectDataLeaks so the explanation on
// screen cannot drift away from the rule that actually fired.
const DOWNLOAD_BURST_HINT = '10 or more files inside 10 minutes';

const DataLeak = ({ report, error, onRefresh, busy }) => {
  const { summary, incidents } = report;
  const [sev, setSev] = useState('all');
  const [query, setQuery] = useState('');

  const visible = useMemo(() => {
    const needle = query.trim().toLowerCase();
    return incidents.filter(i =>
      (sev === 'all' || i.severity === sev)
      && (!needle
        || (i.user || '').toLowerCase().includes(needle)
        || (i.file || '').toLowerCase().includes(needle)
        || (i.type || '').toLowerCase().includes(needle)));
  }, [incidents, sev, query]);

  // Severity counts come from the summary, which is computed over the whole
  // dataset - so a chip can legitimately read higher than the capped table.
  const FILTERS = [
    { key: 'all', label: `All (${summary.totalIncidents ?? incidents.length})` },
    { key: 'critical', label: `Critical (${summary.critical})` },
    { key: 'high', label: `High (${summary.high})` },
    { key: 'medium', label: `Company-wide (${summary.medium})` },
  ];

  const truncated = (summary.totalIncidents ?? 0) > incidents.length;

  return (
    <div className="fadeIn">
      <div className="tile-row">
        <Tile tone="blush" icon={AlertTriangle} label="Critical" value={summary.critical} note="Reachable outside the tenant" />
        <Tile icon={ShieldAlert} label="High risk" value={summary.high} note="Sensitive files and bulk pulls" />
        <Tile tone="sky" icon={Share2} label="Company-wide" value={summary.medium} note="Access widened internally" />
        <Tile tone="mint" icon={FileText} label="Sensitive files" value={summary.sensitiveFiles} note="Matched by filename" />
      </div>

      <div className="chart-panel">
        <PanelHead icon={ShieldX} title="Data Exfiltration Signals" onRefresh={onRefresh} busy={busy} />
        <p className="panel-caption">
          {summary.eventsScanned.toLocaleString()} events scanned from the last 24 hours. Each signal below is
          an observed audit record, not a prediction: a link scoped to Anyone needs no sign-in, a Guest target is
          an account outside this tenant, and a bulk download is {' '}
          {DOWNLOAD_BURST_HINT} — worth a look, not automatically malicious.
        </p>

        {incidents.length > 0 && (
          <div className="filter-row">
            {FILTERS.map(f => (
              <button
                key={f.key}
                className={`filter-chip ${sev === f.key ? 'active' : ''}`}
                onClick={() => setSev(f.key)}
              >
                {f.label}
              </button>
            ))}
            <input
              className="leak-search"
              type="search"
              placeholder="Filter by user, file or signal…"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
        )}

        {truncated && (
          <p className="leak-truncation">
            Showing the top {incidents.length} of {summary.totalIncidents} signals, most severe first.
            Narrow the window or resolve the criticals to see the rest.
          </p>
        )}

        {error ? (
          <div className="data-empty">{error}</div>
        ) : incidents.length === 0 ? (
          <div className="data-empty">No exposure signals in this window.</div>
        ) : visible.length === 0 ? (
          <div className="data-empty">No signals match this filter.</div>
        ) : (
          <div className="table-scroll">
            <table className="custom-table">
              <thead>
                <tr>
                  <th>Severity</th>
                  <th>Signal</th>
                  <th>Identity</th>
                  <th>File</th>
                  <th>What happened</th>
                  <th>Device</th>
                  <th>When</th>
                </tr>
              </thead>
              <tbody>
                {visible.map(incident => (
                  <tr key={incident.id}>
                    <td><span className={`pill ${LEAK_PILL[incident.severity] || 'pill-gray'}`}>{incident.severity.toUpperCase()}</span></td>
                    <td style={{ fontWeight: 700, color: 'var(--text-primary)' }}>{incident.type}</td>
                    <td>{incident.user?.split('@')[0]}<div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{incident.user}</div></td>
                    <td style={{ maxWidth: 240, overflowWrap: 'anywhere' }} title={incident.path}>{incident.file}</td>
                    <td style={{ color: 'var(--text-secondary)', fontSize: '0.75rem', maxWidth: 280 }}>{incident.detail}</td>
                    <td>
                      <span className={`pill ${incident.managedDevice ? 'pill-green' : 'pill-orange'}`}>
                        {incident.managedDevice ? 'Managed' : 'Unmanaged'}
                      </span>
                    </td>
                    <td style={{ color: 'var(--text-secondary)', fontSize: '0.75rem' }}>{formatWhen(incident.when)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

const WebUsage = ({ report, error, onRefresh, busy }) => {
  const { available, reason, totals, users, sites } = report;
  const peak = Math.max(...sites.map(s => s.visits), 1);

  return (
    <div className="fadeIn">
      {available && (
        <div className="tile-row">
          <Tile icon={Globe} label="Page visits" value={totals.visits.toLocaleString()} note="Reported by the desktop agent" />
          <Tile tone="sky" icon={Users} label="People" value={totals.users} note="Accounts seen browsing" />
          <Tile tone="mint" icon={Database} label="Distinct sites" value={totals.sites} note="Unique hostnames" />
          <Tile tone="blush" icon={ShieldAlert} label="Flagged visits" value={totals.flagged} tag={shareOf(totals.flagged, totals.visits)} note="Outside the allowed categories" used={totals.flagged} total={totals.visits} />
        </div>
      )}

      <div className="chart-panel" style={{ marginBottom: 16 }}>
        <PanelHead icon={Globe} title="Most Visited Sites" onRefresh={onRefresh} busy={busy} />
        {error ? (
          <div className="data-empty">{error}</div>
        ) : !available ? (
          // Naming the cause matters more than an empty table: browsing data is
          // absent because nothing is collecting it, not because nobody browses.
          <div className="data-empty">
            <strong style={{ display: 'block', marginBottom: 8, color: 'var(--ink)' }}>No browsing data yet</strong>
            {reason}
          </div>
        ) : (
          <div className="sev-list">
            {sites.map(site => (
              <div className="sev-row" key={site.host}>
                <span className="sev-name" title={site.host}>{site.host}</span>
                <div className="sev-track">
                  <div className="sev-fill" style={{ width: `${(site.visits / peak) * 100}%`, background: 'var(--fill-violet)' }} />
                </div>
                <span className="sev-count">{site.visits}</span>
              </div>
            ))}
          </div>
        )}
      </div>

      {available && (
        <div className="chart-panel">
          <div className="chart-title">Per-Person Breakdown</div>
          <p className="panel-caption">Admin and service accounts are excluded, so this lists staff activity only.</p>
          <div className="table-scroll">
            <table className="custom-table">
              <thead>
                <tr>
                  <th>Person</th>
                  <th>Device</th>
                  <th>Visits</th>
                  <th>Most visited</th>
                  <th>Flagged categories</th>
                  <th>Last seen</th>
                </tr>
              </thead>
              <tbody>
                {users.map(person => (
                  <tr key={person.user}>
                    <td style={{ fontWeight: 700 }}>{person.user}</td>
                    <td style={{ color: 'var(--ink-soft)', fontSize: '0.75rem' }}>{person.devices.join(', ') || '—'}</td>
                    <td className="num">{person.visits}</td>
                    <td>
                      {person.topSites.map(s => (
                        <div key={s.host} style={{ fontSize: '0.75rem' }}>
                          {s.host} <span style={{ color: 'var(--ink-faint)' }}>({s.visits})</span>
                        </div>
                      ))}
                    </td>
                    <td>
                      {person.flaggedCategories.length === 0
                        ? <span className="pill pill-green">Clean</span>
                        : person.flaggedCategories.map(c => (
                            <div key={c.name} style={{ marginBottom: 3 }}>
                              <span className="pill pill-yellow">{c.name} · {c.visits}</span>
                            </div>
                          ))}
                    </td>
                    <td style={{ color: 'var(--ink-soft)', fontSize: '0.75rem' }}>{formatWhen(person.lastSeen)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};

const DeviceFleet = ({ devices, error, onRefresh, busy, API_BASE, canReboot, getIdToken }) => {
  const [filter, setFilter] = useState('all');
  // A reboot lands on somebody's actual machine, so each row tracks its own
  // in-flight state - one shared flag would grey out the whole column and leave
  // the operator unsure which device the command went to.
  const [rebooting, setRebooting] = useState(null);
  const [rebootNote, setRebootNote] = useState(null);

  const handleReboot = async (device) => {
    const name = device.deviceName || 'this device';
    const owner = device.userDisplayName ? ` (${device.userDisplayName})` : '';
    if (!window.confirm(
      `Restart ${name}${owner} now?\n\n`
      + `Intune delivers the command on the device's next check-in. `
      + `Unsaved work on that machine will be lost.`
    )) return;

    setRebooting(device.id);
    setRebootNote(null);
    try {
      // The backend verifies this token against Entra before it will touch any
      // hardware, so it has to be a live one from MSAL - not the copy cached on
      // the account object, which can outlive its expiry.
      const idToken = await getIdToken();
      const res = await fetch(`${API_BASE}/api/device/${device.id}/reboot`, {
        method: 'POST',
        headers: {
          'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$',
          'Authorization': `Bearer ${idToken}`,
        },
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(body.error || `Reboot API responded ${res.status}`);
      setRebootNote({ ok: true, text: `Restart command queued for ${name}. It runs on the next Intune check-in.` });
    } catch (e) {
      setRebootNote({ ok: false, text: `${name}: ${e.message}` });
    } finally {
      setRebooting(null);
    }
  };

  const compliant = devices.filter(d => d.complianceState === 'compliant').length;
  const nonCompliant = devices.filter(d => d.complianceState === 'noncompliant').length;
  // Intune also reports states like "unknown" and "inGracePeriod"; counting those
  // as compliant would flatter the fleet, so they get their own bucket.
  const otherState = devices.length - compliant - nonCompliant;

  // Devices needing attention sort to the top - a fleet table ordered by whatever
  // Graph returned makes the operator hunt for the ones that matter.
  const visible = useMemo(() => {
    const rank = { noncompliant: 0, unknown: 1, inGracePeriod: 1, compliant: 3 };
    return devices
      .filter(d => filter === 'all'
        || (filter === 'compliant' && d.complianceState === 'compliant')
        || (filter === 'noncompliant' && d.complianceState === 'noncompliant')
        || (filter === 'other' && !['compliant', 'noncompliant'].includes(d.complianceState)))
      .slice()
      .sort((a, b) => (rank[a.complianceState] ?? 2) - (rank[b.complianceState] ?? 2)
        || String(a.deviceName).localeCompare(String(b.deviceName)));
  }, [devices, filter]);

  const FILTERS = [
    { key: 'all', label: `All (${devices.length})` },
    { key: 'noncompliant', label: `Not compliant (${nonCompliant})` },
    { key: 'other', label: `Unknown / grace (${otherState})` },
    { key: 'compliant', label: `Compliant (${compliant})` },
  ];

  return (
    <div className="fadeIn">
      {devices.length > 0 && (
        <div className="tile-row">
          <Tile icon={Laptop} label="Managed" value={devices.length} note="Enrolled in Intune" />
          <Tile tone="blush" icon={AlertTriangle} label="Not compliant" value={nonCompliant} tag={shareOf(nonCompliant, devices.length)} note="Failing policy right now" used={nonCompliant} total={devices.length} />
          <Tile tone="sky" icon={ShieldAlert} label="Unknown" value={otherState} tag={shareOf(otherState, devices.length)} note="No verdict yet" used={otherState} total={devices.length} />
          <Tile tone="mint" icon={ShieldCheck} label="Compliant" value={compliant} tag={shareOf(compliant, devices.length)} note="Meeting policy" used={compliant} total={devices.length} />
        </div>
      )}

      <div className="chart-panel">
        <PanelHead icon={Laptop} title="Corporate Device Fleet" onRefresh={onRefresh} busy={busy} />

        {devices.length > 0 && (
          <div className="filter-row">
            {FILTERS.map(f => (
              <button
                key={f.key}
                className={`filter-chip ${filter === f.key ? 'active' : ''}`}
                onClick={() => setFilter(f.key)}
              >
                {f.label}
              </button>
            ))}
          </div>
        )}

        {rebootNote && (
          <div className={`reboot-note ${rebootNote.ok ? 'ok' : 'bad'}`}>
            <span>{rebootNote.text}</span>
            <button onClick={() => setRebootNote(null)} aria-label="Dismiss">×</button>
          </div>
        )}

        {error ? (
          <div className="data-empty">{error}</div>
        ) : devices.length === 0 ? (
          <div className="data-empty">No managed devices returned by Intune.</div>
        ) : visible.length === 0 ? (
          <div className="data-empty">No devices in this state.</div>
        ) : (
          <div className="table-scroll">
            <table className="custom-table">
              <thead>
                <tr>
                  <th>Device</th>
                  <th>Assigned to</th>
                  <th>Platform</th>
                  <th>Compliance</th>
                  <th>Last check-in</th>
                  {canReboot && <th>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {visible.map(device => (
                  <tr key={device.id}>
                    <td>
                      <div style={{ fontWeight: 700, color: 'var(--text-primary)' }}>{device.deviceName || 'Unnamed device'}</div>
                      <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{device.model || '—'}</div>
                    </td>
                    <td>
                      <div>{device.userDisplayName || 'Unassigned'}</div>
                      <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{device.userPrincipalName || ''}</div>
                    </td>
                    <td>{device.operatingSystem || '—'}</td>
                    <td>
                      <span className={`pill ${device.complianceState === 'compliant' ? 'pill-green' : device.complianceState === 'noncompliant' ? 'pill-red' : 'pill-gray'}`}>
                        {device.complianceState || 'unknown'}
                      </span>
                    </td>
                    <td style={{ color: 'var(--text-secondary)', fontSize: '0.75rem' }}>{formatWhen(device.lastSyncDateTime)}</td>
                    {canReboot && <td>
                      <button
                        className="row-action danger"
                        onClick={() => handleReboot(device)}
                        disabled={rebooting !== null || !device.id}
                        title={device.id ? `Send a restart command to ${device.deviceName || 'this device'}` : 'Intune returned no device id'}
                      >
                        {rebooting === device.id
                          ? <><RefreshCw size={12} className="spinning" /> Sending…</>
                          : <><Power size={12} /> Restart</>}
                      </button>
                    </td>}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

const AdminPortal = ({ authorizedUsers, fetchAuthList, API_BASE }) => {
  const [newEmail, setNewEmail] = useState('');
  const [loading, setLoading] = useState(false);
  const handleAdd = async () => {
    if (!newEmail.includes('@')) return alert("Invalid Email");
    setLoading(true);
    try {
      await fetch(`${API_BASE}/api/admin/authorized-users`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' },
        body: JSON.stringify({ email: newEmail.toLowerCase() })
      });
      setNewEmail('');
      await fetchAuthList();
    } finally { setLoading(false); }
  };
  const handleRemove = async (email) => {
    if (!window.confirm(`Revoke access for ${email}?`)) return;
    setLoading(true);
    try {
      await fetch(`${API_BASE}/api/admin/authorized-users`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json', 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' },
        body: JSON.stringify({ email: email.toLowerCase() })
      });
      await fetchAuthList();
    } finally { setLoading(false); }
  };
  return (
    <div className="admin-portal-container fadeIn">
      <div className="chart-panel" style={{ maxWidth: '800px', margin: '0 auto', padding: '30px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px', marginBottom: '30px' }}>
          <ShieldCheck size={32} color="var(--accent-primary)" />
          <h2 style={{ margin: 0, fontSize: '1.5rem' }}>Access Control Terminal</h2>
        </div>
        <div style={{ display: 'flex', gap: '10px', marginBottom: '40px' }}>
          <input type="email" placeholder="Enter Corporate Email..." value={newEmail} onChange={(e) => setNewEmail(e.target.value)} style={{ flex: 1, padding: '12px', background: 'var(--panel-bg)', border: '1px solid var(--border-color)', borderRadius: '8px', color: '#fff' }} />
          <button onClick={handleAdd} disabled={loading} className="btn-primary" style={{ background: 'var(--accent-primary)', color: '#000', border: 'none', padding: '0 20px', borderRadius: '8px', fontWeight: '700' }}>Grant Access</button>
        </div>
        <div className="users-auth-list">
          {authorizedUsers.map(email => (
            <div key={email} className="feed-item" style={{ marginBottom: '10px' }}>
              <span>{email}</span>
              <button onClick={() => handleRemove(email)} style={{ marginLeft: 'auto', background: 'transparent', border: 'none', color: 'var(--critical)' }}><Trash2 size={18} /></button>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

const SUPER_ADMINS = ['help-desk@ldplogistics.com', 'kundan@ldplogistics.com'];

function App() {
  const { instance, accounts } = useMsal();
  const [authorizedUsers, setAuthorizedUsers] = useState([]);
  const [username, setUsername] = useState('help-desk@ldplogistics.com');
  const [pwd, setPwd] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [authLoading, setAuthLoading] = useState(true);
  const [adminStatus, setAdminStatus] = useState({ isSuperAdmin: false, isAuthorized: false });
  const [authError, setAuthError] = useState(null);
  const [lastSync, setLastSync] = useState(null);
  const [clockTick, setClockTick] = useState(() => Date.now());
  const [data, setData] = useState([]);
  const [error, setError] = useState(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [theme, setTheme] = useState('studio'); // 'studio' (light) or 'dark'
  
  // Dev: Vite proxies /api to localhost:3001. Prod: VITE_API_URL is baked in at
  // build time; if that build arg is ever missing, fall back to the "api"-prefixed
  // sibling host (ldpm365.ldplogistics.com -> apildpm365.ldplogistics.com) so the
  // dashboard never silently fetches HTML from its own origin.
  const API_BASE = useMemo(() => {
    if (import.meta.env.DEV) return '';
    const configured = import.meta.env.VITE_API_URL;
    if (configured) return configured.replace(/\/+$/, '');
    const { protocol, host, hostname } = window.location;
    const servedLocally = hostname === 'localhost' || hostname === '127.0.0.1' ||
      /^\d+\.\d+\.\d+\.\d+$/.test(hostname);
    return servedLocally ? '' : `${protocol}//api${host}`;
  }, []);

  // Theme-aware Chart Colors
  // Chart libraries paint onto a canvas and cannot resolve CSS custom properties,
  // so the marks need literal values even though the surrounding chrome is themed.
  const activeColors = useMemo(() => (theme === 'dark' ? {
    accent: '#a5a6f6',
    accentFill: 'rgba(165, 166, 246, 0.18)',
    bar: '#7f80e0',
    grid: 'rgba(255, 255, 255, 0.07)',
    text: '#8a8b92',
  } : {
    accent: '#8b8cf0',
    accentFill: 'rgba(165, 166, 246, 0.22)',
    bar: '#c3d4e2',
    grid: '#f0f0f2',
    text: '#8a8b92',
  }), [theme]);

  const metrics = useMemo(() => {
    const users = new Set(data.map(d => d.UserId).filter(Boolean));
    const count = (re) => data.filter(d => re.test(d.Operation || '')).length;
    return {
      total: data.length,
      users: users.size,
      deletions: count(/delete|purge/i),
      shares: count(/shar|anonymouslink|companylink/i),
      failedLogins: data.filter(d => /failed/i.test(d.Operation || '') || d.ResultStatus === 'Failed').length,
    };
  }, [data]);

  const severityBreakdown = useMemo(() => {
    const counts = {};
    data.forEach(d => {
      const key = severityOf(d.Operation);
      counts[key] = (counts[key] || 0) + 1;
    });
    return SEVERITY_RULES.map(rule => ({ ...rule, count: counts[rule.key] || 0 }));
  }, [data]);

  const topOperations = useMemo(() => {
    const counts = {};
    data.forEach(d => { if (d.Operation) counts[d.Operation] = (counts[d.Operation] || 0) + 1; });
    return Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6)
      .map(([name, count]) => ({
        name,
        count,
        color: SEVERITY_RULES.find(r => r.key === severityOf(name))?.color,
      }));
  }, [data]);

  // Twelve fixed hourly buckets so the axis stays put between refreshes; empty
  // hours are kept at zero rather than skipped, otherwise a quiet night reads as
  // a continuous busy line.
  const activityTimeline = useMemo(() => {
    if (!data.length) return [];
    const buckets = new Map();
    const cursor = new Date();
    cursor.setMinutes(0, 0, 0);
    for (let i = 11; i >= 0; i--) buckets.set(cursor.getTime() - i * 3600000, 0);

    data.forEach(entry => {
      const at = new Date(entry.CreationTime);
      if (Number.isNaN(at.getTime())) return;
      at.setMinutes(0, 0, 0);
      const key = at.getTime();
      if (buckets.has(key)) buckets.set(key, buckets.get(key) + 1);
    });

    return [...buckets].map(([time, count]) => ({
      hour: new Date(time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      count,
    }));
  }, [data]);

  const entitiesData = useMemo(() => {
    const userCounts = {};
    data.forEach(d => { if (d.UserId) userCounts[d.UserId] = (userCounts[d.UserId] || 0) + 1; });
    return Object.entries(userCounts).sort((a,b) => b[1] - a[1]).slice(0, 5).map(([user, count]) => ({ name: user.split('@')[0], count }));
  }, [data]);

  // Privileged endpoints verify the caller's Microsoft identity server-side, so
  // they need a live ID token. MSAL renews it silently; if the session has aged
  // out past silent renewal, a popup is the only way back without losing the page.
  const getIdToken = async () => {
    const account = accounts[0];
    if (!account) throw new Error('No signed-in account');
    try {
      const result = await instance.acquireTokenSilent({ ...loginRequest, account });
      return result.idToken;
    } catch {
      const result = await instance.acquireTokenPopup({ ...loginRequest, account });
      return result.idToken;
    }
  };

  const fetchAuthList = async () => {
    const email = accounts[0]?.username?.toLowerCase();
    try {
      const response = await fetch(`${API_BASE}/api/admin/authorized-users`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (!response.ok) throw new Error(`API responded ${response.status}`);
      const list = await response.json();
      const authorizedList = Array.isArray(list) ? list : [];
      setAuthorizedUsers(authorizedList);
      // Membership of the Admin Portal list is the only thing that grants
      // access. Super admin decides who may edit that list, not who gets in -
      // a super admin still has to appear in it, and the API keeps them there.
      setAdminStatus({
        isSuperAdmin: SUPER_ADMINS.includes(email),
        isAuthorized: authorizedList.includes(email),
      });
      setAuthError(null);
    } catch (e) {
      console.error("Auth sync failed", e);
      // An unreachable API is not a rejected identity. Report the outage rather
      // than claiming the account lacks access, which sends people hunting for
      // a permissions problem that isn't there.
      setAdminStatus({ isSuperAdmin: false, isAuthorized: false });
      setAuthError(e.message);
    } finally { setAuthLoading(false); }
  };

  const fetchData = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/audit-logs`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (!response.ok) throw new Error(`API responded ${response.status}`);
      setData(await response.json());
      setLastSync(Date.now());
      setError(null);
    } catch (err) { setError(err.message); }
  };

  const [riskStats, setRiskStats] = useState([]);
  const [devices, setDevices] = useState([]);
  const [devicesError, setDevicesError] = useState(null);
  const [leakReport, setLeakReport] = useState({ summary: { critical: 0, high: 0, medium: 0, sensitiveFiles: 0, eventsScanned: 0, totalIncidents: 0, shown: 0 }, incidents: [] });
  const [leakError, setLeakError] = useState(null);
  const [webReport, setWebReport] = useState({ available: false, reason: null, totals: { visits: 0, users: 0, sites: 0, flagged: 0 }, users: [], sites: [] });
  const [webError, setWebError] = useState(null);
  const [refreshing, setRefreshing] = useState(null);

  // Wraps a fetch so the panel's own button can show it working. Keyed by panel
  // so one refresh does not spin every other button on screen.
  const manualRefresh = (key, ...fetchers) => async () => {
    setRefreshing(key);
    try { await Promise.all(fetchers.map(fn => fn())); } finally { setRefreshing(null); }
  };

  const fetchRiskStats = async () => {
    try {
      const res = await fetch(`${API_BASE}/api/security/risk-stats`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (res.ok) setRiskStats(await res.json());
    } catch (e) { console.error('Risk stats fetch failed', e); }
  };

  const fetchDataLeak = async () => {
    try {
      const res = await fetch(`${API_BASE}/api/security/data-leak`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (!res.ok) throw new Error(`Scan responded ${res.status}`);
      setLeakReport(await res.json());
      setLeakError(null);
    } catch (e) { setLeakError(e.message); }
  };

  const fetchWebUsage = async () => {
    try {
      const res = await fetch(`${API_BASE}/api/security/browsing`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (!res.ok) throw new Error(`Browsing API responded ${res.status}`);
      setWebReport(await res.json());
      setWebError(null);
    } catch (e) { setWebError(e.message); }
  };

  const fetchDevices = async () => {
    try {
      const res = await fetch(`${API_BASE}/api/devices/all`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      // Intune returns an error body rather than an empty list when the tenant has
      // no managed devices, so show that message instead of an empty table.
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.error || `Device API responded ${res.status}`);
      }
      setDevices(await res.json());
      setDevicesError(null);
    } catch (e) { setDevicesError(e.message); }
  };

  const userIntelligenceSummary = useMemo(() => {
    if (!data || data.length === 0) return [];
    const stats = {};
    data.forEach(log => {
        const userId = log.UserId?.toLowerCase();
        if (!userId || userId.includes('app@sharepoint')) return;
        if (!stats[userId]) {
            stats[userId] = { id: userId, actions: 0, lastSeen: log.CreationTime, workload: log.Workload };
        }
        stats[userId].actions++;
        if (new Date(log.CreationTime) > new Date(stats[userId].lastSeen)) stats[userId].lastSeen = log.CreationTime;
    });
    const now = new Date();
    return Object.values(stats).map(u => ({ ...u, isLive: (now - new Date(u.lastSeen)) < (15 * 60 * 1000) })).sort((a,b) => b.actions - a.actions);
  }, [data]);

  useEffect(() => { 
    if (accounts.length > 0) {
      fetchAuthList(); 
      const timeout = setTimeout(() => setAuthLoading(false), 8000);
      return () => clearTimeout(timeout);
    } 
  }, [accounts]);

  // activeTab is deliberately NOT a dependency here. It used to be, which tore the
  // timer down and restarted the countdown on every tab click - so someone moving
  // around the dashboard could sit indefinitely without a single refresh landing.
  useEffect(() => {
    if (!accounts.length || !adminStatus.isAuthorized) return;
    // The leak scan runs on data the backend already holds - no Graph call - so it
    // rides the main interval. That keeps the sidebar's critical badge live
    // whichever tab is open; a leak you only find by clicking the tab is a leak
    // nobody finds.
    const poll = () => { fetchData(); fetchDataLeak(); };
    poll();
    const interval = setInterval(poll, REFRESH_MS);
    return () => clearInterval(interval);
  }, [accounts.length, adminStatus.isAuthorized]);

  // Tab-scoped data hits Microsoft Graph on every call, so it refreshes when its
  // tab is opened and then at a slower cadence than the audit feed.
  useEffect(() => {
    if (!accounts.length || !adminStatus.isAuthorized) return;
    const load = () => {
      if (activeTab === 'behavior') fetchRiskStats();
      if (activeTab === 'devices') fetchDevices();
      if (activeTab === 'web') fetchWebUsage();
    };
    load();
    const interval = setInterval(load, REFRESH_MS * 3);
    return () => clearInterval(interval);
  }, [accounts.length, adminStatus.isAuthorized, activeTab]);

  // Drives the "synced Ns ago" readout so it keeps counting between fetches.
  useEffect(() => {
    const tick = setInterval(() => setClockTick(Date.now()), 1000);
    return () => clearInterval(tick);
  }, []);

  const syncAge = lastSync ? Math.round((clockTick - lastSync) / 1000) : null;
  const syncState = error ? 'down' : syncAge !== null && syncAge > (REFRESH_MS / 1000) * 3 ? 'stale' : 'ok';

  const handleLogin = () => instance.loginRedirect(loginRequest);

  return (
    <div className={`studio ${theme === 'dark' ? 'studio-dark' : ''}`}>
      <AuthenticatedTemplate>
        {authLoading ? (
          <div className="loading-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100vh', width: '100vw', background: 'var(--main-bg)', color: 'var(--accent-primary)' }}>
            <div className="loading-spinner"></div>
            <p style={{ fontWeight: 800, marginTop: '20px' }}>VERIFYING TERMINAL ACCESS...</p>
          </div>
        ) : !adminStatus.isAuthorized ? (
          <div className="access-denied-wrapper fadeIn" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', width: '100vw', background: 'var(--main-bg)' }}>
            <div className="chart-panel" style={{ textAlign: 'center' }}>
              <h1 style={{ color: 'var(--critical)' }}>{authError ? 'SERVICE UNAVAILABLE' : 'ACCESS RESTRICTED'}</h1>
              <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                {authError
                  ? `Could not reach the security API (${authError}). This is a connectivity problem, not a permissions one.`
                  : 'Your identity has not been authorized for terminal access.'}
              </p>
              {accounts[0]?.username && (
                <p style={{ color: 'var(--text-muted)', fontSize: '0.75rem', marginBottom: '20px' }}>
                  Signed in as {accounts[0].username}
                </p>
              )}
              <button onClick={() => instance.logoutRedirect()} className="btn-primary">Logout</button>
            </div>
          </div>
        ) : (
          <>
            <aside className="icon-rail">
              <div className="rail-mark">LDP</div>
              {TABS.filter(t => t.key !== 'admin' || adminStatus.isSuperAdmin).map(tab => (
                <button
                  key={tab.key}
                  className={`rail-btn ${activeTab === tab.key ? 'active' : ''}`}
                  onClick={() => setActiveTab(tab.key)}
                  title={tab.label}
                  aria-label={tab.label}
                >
                  <tab.icon size={19} />
                  {tab.key === 'leak' && leakReport.summary.critical > 0 && (
                    <span className="rail-dot">{leakReport.summary.critical}</span>
                  )}
                </button>
              ))}
              <div className="rail-spacer" />
              <button
                className="rail-avatar"
                onClick={() => instance.logoutRedirect()}
                title={`${accounts[0]?.username || ''} — sign out`}
              >
                {(accounts[0]?.name || 'U').slice(0, 1).toUpperCase()}
              </button>
            </aside>

            <main className="workspace">
              <div className="workspace-head">
                <div className="brand-block">
                  {/* Wordmark is drawn in type rather than shipped as an image so it
                      stays crisp at any size and follows the theme's ink colour. */}
                  <h1 className="brand" aria-label="LDP Logistics">
                    <span className="brand-ldp">LDP</span>
                    <span className="brand-name">Logistics</span>
                    <span className="brand-dot" aria-hidden="true" />
                  </h1>
                  <p className="brand-sub">Security Dashboard</p>
                </div>
                <div className="head-actions">
                  {/* Reports what the last fetch actually did. The old header claimed
                      SECURE CONNECTION ACTIVE unconditionally, so a dead API and a
                      healthy one looked identical. */}
                  <span className={`sync-chip ${syncState}`}>
                    <span className="dot" />
                    {syncState === 'down'
                      ? `Sync failed — ${error}`
                      : syncAge === null
                        ? 'Connecting…'
                        : `Live · synced ${syncAge}s ago`}
                  </span>
                  <button className="ghost-btn" onClick={() => setTheme(theme === 'studio' ? 'dark' : 'studio')}>
                    <Settings size={15} /> {theme === 'studio' ? 'Dark' : 'Light'}
                  </button>
                  <button className="dark-btn" onClick={() => setActiveTab('leak')}>
                    <ShieldX size={15} /> Review exposure
                  </button>
                </div>
              </div>

              <nav className="pill-tabs">
                {TABS.filter(t => t.key !== 'admin' || adminStatus.isSuperAdmin).map(tab => (
                  <button
                    key={tab.key}
                    className={`pill-tab ${activeTab === tab.key ? 'active' : ''}`}
                    onClick={() => setActiveTab(tab.key)}
                  >
                    {tab.label}
                    {tab.key === 'leak' && leakReport.summary.critical > 0 && (
                      <span className="tab-count">{leakReport.summary.critical}</span>
                    )}
                  </button>
                ))}
              </nav>

              <div className="workspace-body">
                {activeTab === 'dashboard' && (
                  <DashboardContent
                    data={data}
                    metrics={metrics}
                    severityBreakdown={severityBreakdown}
                    topOperations={topOperations}
                    activityTimeline={activityTimeline}
                    entitiesData={entitiesData}
                    chartColors={activeColors}
                    onRefresh={manualRefresh('dashboard', fetchData, fetchDataLeak)}
                    busy={refreshing === 'dashboard'}
                  />
                )}
                
                {activeTab === 'users' && (
                  <div className="fadeIn">
                    <div className="chart-panel">
                      <div className="panel-head">
                        <div className="chart-title" style={{ marginBottom: 0 }}>Employee Intelligence Overview</div>
                        <RefreshButton onRefresh={manualRefresh('users', fetchData, fetchAuthList)} busy={refreshing === 'users'} />
                      </div>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th>Identity</th>
                            <th>Status</th>
                            <th>Risk Level</th>
                            <th>Activity Count</th>
                            <th>Dashboard Access</th>
                            <th>Investigate</th>
                          </tr>
                        </thead>
                        <tbody>
                          {userIntelligenceSummary.map(u => (
                            <tr key={u.id}>
                              <td><div style={{ fontWeight: 700, color: 'var(--text-primary)' }}>{u.id.split('@')[0]}</div><div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{u.id}</div></td>
                              <td><span className={`pill ${u.isLive ? 'pill-cyan' : 'pill-gray'}`}>{u.isLive ? 'LIVE' : 'IDLE'}</span></td>
                              <td><span className={`pill ${u.actions > 50 ? 'pill-red' : 'pill-cyan'}`}>{u.actions > 50 ? 'CRITICAL' : 'LOW'}</span></td>
                              <td>{u.actions} hits</td>
                              <td>
                                {adminStatus.isSuperAdmin ? (
                                  authorizedUsers.includes(u.id.toLowerCase()) ? (
                                    <button
                                      style={{ padding: '4px 12px', background: 'rgba(255,77,79,0.1)', color: '#ff4d4f', border: '1px solid rgba(255,77,79,0.2)', borderRadius: '4px', cursor: 'pointer', fontSize: '0.7rem', display: 'flex', alignItems: 'center', gap: '5px' }}
                                      onClick={async () => {
                                        if (!window.confirm('Revoke access for ' + u.id + '?')) return;
                                        await fetch(API_BASE + '/api/admin/authorized-users', {
                                          method: 'DELETE',
                                          headers: { 'Content-Type': 'application/json', 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' },
                                          body: JSON.stringify({ email: u.id.toLowerCase() })
                                        });
                                        await fetchAuthList();
                                      }}
                                    >
                                      <ShieldX size={12} /> Revoke
                                    </button>
                                  ) : (
                                    <button
                                      style={{ padding: '4px 12px', background: 'rgba(0,184,148,0.1)', color: 'var(--accent-primary)', border: '1px solid rgba(0,184,148,0.2)', borderRadius: '4px', cursor: 'pointer', fontSize: '0.7rem', display: 'flex', alignItems: 'center', gap: '5px' }}
                                      onClick={async () => {
                                        await fetch(API_BASE + '/api/admin/authorized-users', {
                                          method: 'POST',
                                          headers: { 'Content-Type': 'application/json', 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' },
                                          body: JSON.stringify({ email: u.id.toLowerCase() })
                                        });
                                        await fetchAuthList();
                                      }}
                                    >
                                      <ShieldCheck size={12} /> Grant Access
                                    </button>
                                  )
                                ) : (
                                  <span style={{ fontSize: '0.7rem', color: authorizedUsers.includes(u.id.toLowerCase()) ? 'var(--success)' : 'var(--text-muted)' }}>
                                    {authorizedUsers.includes(u.id.toLowerCase()) ? '✓ Authorized' : '✗ Restricted'}
                                  </span>
                                )}
                              </td>
                              <td><button className="btn-primary" 
                                          style={{ padding: '4px 12px', borderRadius: '4px', background: 'rgba(59, 130, 246, 0.1)', color: 'var(--accent-primary)', border: '1px solid var(--accent-primary)' }}
                                          onClick={() => setActiveTab('dashboard')}
                                  >Investigate</button></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {activeTab === 'admin' && <AdminPortal authorizedUsers={authorizedUsers} fetchAuthList={fetchAuthList} API_BASE={API_BASE} />}
                
                {activeTab === 'behavior' && <SecurityPulse riskStats={riskStats} onRefresh={manualRefresh("behavior", fetchRiskStats)} busy={refreshing === "behavior"} />}

                {activeTab === 'devices' && <DeviceFleet devices={devices} error={devicesError} onRefresh={manualRefresh("devices", fetchDevices)} busy={refreshing === "devices"} API_BASE={API_BASE} canReboot={adminStatus.isSuperAdmin} getIdToken={getIdToken} />}

                {activeTab === 'leak' && <DataLeak report={leakReport} error={leakError} onRefresh={manualRefresh("leak", fetchDataLeak)} busy={refreshing === "leak"} />}

                {activeTab === 'web' && <WebUsage report={webReport} error={webError} onRefresh={manualRefresh('web', fetchWebUsage)} busy={refreshing === 'web'} />}
              </div>
            </main>
          </>
        )}
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <div className="dark-cyber-login-container">
          <div className="cyber-scanline"></div>

          {/* Logo Brand Overlay */}
          <div className="cyber-logo-area">
            <span className="logo-bold">LDP</span>
            <span className="logo-light">Logistics</span>
            <span className="logo-dot">.</span>
          </div>

          {/* Left Panel - Hero Section */}
          <div className="cyber-hero-panel">
            <div className="cyber-hero-badge">
              <div className="cyber-hero-badge-dot"></div>
              CyberSec Intelligence Platform
            </div>

            <h1 className="cyber-hero-title">
              Secure Your Digital Environment with Real-Time Cyber Intelligence
            </h1>

            <p className="cyber-hero-desc">
              Monitor security events, identify vulnerabilities, detect cyber threats, and gain complete visibility into your organization's security posture through one intelligent dashboard.
            </p>

            <div className="cyber-features-grid">
              <div className="cyber-feature-item">
                <div className="cyber-feature-icon-wrapper">
                  <ShieldAlert size={18} />
                </div>
                <div className="cyber-feature-text">Real-Time Threat Monitoring</div>
              </div>

              <div className="cyber-feature-item">
                <div className="cyber-feature-icon-wrapper">
                  <Activity size={18} />
                </div>
                <div className="cyber-feature-text">AI-Powered Detection</div>
              </div>

              <div className="cyber-feature-item">
                <div className="cyber-feature-icon-wrapper">
                  <Search size={18} />
                </div>
                <div className="cyber-feature-text">Vulnerability Management</div>
              </div>

              <div className="cyber-feature-item">
                <div className="cyber-feature-icon-wrapper">
                  <Laptop size={18} />
                </div>
                <div className="cyber-feature-text">Endpoint Protection</div>
              </div>

              <div className="cyber-feature-item">
                <div className="cyber-feature-icon-wrapper">
                  <ShieldCheck size={18} />
                </div>
                <div className="cyber-feature-text">Security Analytics</div>
              </div>

              <div className="cyber-feature-item">
                <div className="cyber-feature-icon-wrapper">
                  <FileText size={18} />
                </div>
                <div className="cyber-feature-text">Compliance & Risk Reporting</div>
              </div>
            </div>
          </div>

          {/* Right Panel - Glassmorphic Login Form */}
          <div className="cyber-login-panel">
            <div className="cyber-login-glass-card">
              <h2>Welcome Back</h2>
              <p className="cyber-login-glass-card-tagline">
                Sign in to access your Cybersecurity Dashboard.
              </p>

              <form onSubmit={(e) => { e.preventDefault(); handleLogin(); }}>
                <div className="cyber-input-group">
                  <label htmlFor="username-field">Username</label>
                  <div className="cyber-input-field-wrapper">
                    <input
                      id="username-field"
                      type="email"
                      value={username}
                      onChange={(e) => setUsername(e.target.value)}
                      placeholder="name@example.com"
                    />
                  </div>
                </div>

                <div className="cyber-input-group">
                  <label htmlFor="password-field">Password</label>
                  <div className="cyber-input-field-wrapper">
                    <input
                      id="password-field"
                      type={showPassword ? "text" : "password"}
                      value={pwd}
                      onChange={(e) => setPwd(e.target.value)}
                      placeholder="Enter password"
                    />
                  </div>
                </div>

                <div className="cyber-show-password-wrapper" onClick={() => setShowPassword(!showPassword)}>
                  <input
                    type="checkbox"
                    checked={showPassword}
                    readOnly
                  />
                  <span>Show password</span>
                </div>

                <button type="submit" className="cyber-primary-signin-btn">
                  Sign In
                </button>
              </form>

              <div className="cyber-login-divider">
                <span>Or continue with</span>
              </div>

              <button className="cyber-sso-login-btn" onClick={handleLogin}>
                <div className="microsoft-logo">
                  <div className="microsoft-square red"></div>
                  <div className="microsoft-square green"></div>
                  <div className="microsoft-square blue"></div>
                  <div className="microsoft-square yellow"></div>
                </div>
                Microsoft Single Sign-On
              </button>
            </div>
          </div>
        </div>
      </UnauthenticatedTemplate>
    </div>
  );
}

export default App;
