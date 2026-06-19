// v1.0.3 - Restored clean structure after sync errors
import React, { useState, useEffect, useMemo } from 'react';
import { 
  Activity, Users, FileText, Share2, LogIn, LayoutDashboard, Database, 
  User as UserIcon, ShieldAlert, Laptop, Globe, Lock, Shield, Settings, 
  UserPlus, UserMinus, ShieldCheck, ShieldX, Search, Plus, Trash2, 
  AlertTriangle, Camera, Clock, ShieldHalf, Bell
} from 'lucide-react';
import { 
  Chart as ChartJS, ArcElement, Tooltip as ChartTooltip, Legend as ChartLegend, 
  CategoryScale, LinearScale, BarElement, Title, PointElement, LineElement, Filler
} from 'chart.js';
import { Doughnut, Bar as BarChartJS } from 'react-chartjs-2';
import { ResponsiveSankey } from '@nivo/sankey';
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, BarChart, Bar } from 'recharts';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import { getLiveTime, chartColors, generateActivityFeed } from './dashboard';
import './App.css';

ChartJS.register(ArcElement, ChartTooltip, ChartLegend, CategoryScale, LinearScale, BarElement, Title, PointElement, LineElement, Filler);

const SystemClock = () => {
    const [timeData, setTimeData] = useState(getLiveTime());
    useEffect(() => {
        const timer = setInterval(() => setTimeData(getLiveTime()), 1000);
        return () => clearInterval(timer);
    }, []);
    return (
        <div className="clock-display">
            <span className="clock-time">{timeData.time}</span>
            <span className="clock-date">{timeData.date}</span>
        </div>
    );
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
                <span className="nav-badge" style={{ background: 'var(--accent-cyan)', color: '#000' }}>LIVE</span>
            </div>
            <div className="feed-scroll">
                {feed.map(item => (
                    <div key={item.id} className="feed-item">
                        <span className={`pulse-dot ${item.status === 'alert' ? 'red' : item.status === 'warn' ? 'yellow' : 'green'}`}></span>
                        <div className="feed-info">
                            <div className="feed-user-action">
                                <span style={{ color: 'var(--accent-cyan)' }}>{item.user}</span>
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

const KpiCard = ({ label, value, type, icon: Icon, progress, subtext }) => (
  <div className={`stat-card ${type} fadeIn`}>
    {Icon && <Icon className="stat-icon" />}
    <div className="stat-label">{label}</div>
    <div className="stat-value">{value}</div>
    {progress !== undefined && (
      <div className="progress-bar-container">
        <div className={`progress-bar-fill ${type}`} style={{ width: `${progress}%` }}></div>
      </div>
    )}
    <div className="stat-subtext">{subtext}</div>
  </div>
);

const HeatMapTable = () => {
  const rows = [
    { label: 'Severe', type: 'pill-red' },
    { label: 'Major', type: 'pill-orange' },
    { label: 'Moderate', type: 'pill-yellow' },
    { label: 'Minor', type: 'pill-cyan' },
    { label: 'Insignificance', type: 'pill-gray' }
  ];
  const cols = ['Rare', 'Unlikely', 'Moderate', 'Likely', 'Almost certain'];
  return (
    <div className="fadeIn">
      <table className="custom-table">
        <thead>
          <tr>
            <th style={{ width: '140px' }}>Risk Rating</th>
            {cols.map(c => <th key={c}>{c}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.label}>
              <td style={{ fontWeight: '700', color: '#fff' }}>{r.label}</td>
              {cols.map((c) => (
                <td key={c}>
                  <span className={`pill ${r.type}`}>
                    {Math.floor(Math.random() * 25) + 5}
                  </span>
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

const DashboardContent = React.memo(({ data, metrics, entitiesData, handleUserClick }) => {
  return (
    <div className="fadeIn">
      <div className="grid-metrics">
        <KpiCard label="% Risks >= threshold" value={`${metrics.riskPercent}%`} type="cyan" icon={ShieldAlert} progress={metrics.riskPercent} subtext="Total user risk ratio" />
        <KpiCard label="# Of risks >= threshold" value={metrics.riskyUsers} type="red" icon={AlertTriangle} subtext="Severe risk entities" />
        <KpiCard label="Risks analysis progress" value={`${metrics.progress}%`} type="yellow" icon={Activity} progress={metrics.progress} subtext="Current scan status" />
        <KpiCard label="Response progress" value="52.6%" type="green" icon={ShieldCheck} progress={52.6} subtext="Threat mitigation rate" />
      </div>

      <div className="grid-analysis">
        <div className="chart-panel">
          <div className="chart-title"><ShieldAlert size={18} color="var(--accent-cyan)" /> Risk Rating Breakdown</div>
          <div style={{ height: '220px', position: 'relative' }}>
            <Doughnut 
              data={{
                labels: ['Severe', 'Major', 'Moderate', 'Minor', 'Insignificance'],
                datasets: [{
                  data: [14, 25, 33, 48, 12],
                  backgroundColor: [chartColors.severe, chartColors.major, chartColors.moderate, chartColors.minor, chartColors.insignificance],
                  borderColor: '#0d1629', borderWidth: 2, cutout: '75%'
                }]
              }}
              options={{ maintainAspectRatio: false, plugins: { legend: { display: false } } }}
            />
            <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', textAlign: 'center' }}>
              <div style={{ fontSize: '1.5rem', fontWeight: '800' }}>132</div>
              <div style={{ fontSize: '0.6rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Total Risks</div>
            </div>
          </div>
        </div>

        <div className="chart-panel">
          <div className="chart-title"><Activity size={18} color="var(--accent-cyan)" /> Risk Heart Map</div>
          <HeatMapTable />
        </div>

        <div className="chart-panel">
          <div className="chart-title"><ShieldHalf size={18} color="var(--accent-cyan)" /> Action Plan Breakdown</div>
          <div style={{ height: '180px', position: 'relative' }}>
            <Doughnut 
              data={{
                labels: ['Completed', 'In Progress', 'Pending'],
                datasets: [{
                  data: [35, 45, 20],
                  backgroundColor: [chartColors.completed, chartColors.inProgress, chartColors.pending],
                  borderColor: '#0d1629', borderWidth: 2, cutout: '75%'
                }]
              }}
              options={{ maintainAspectRatio: false, plugins: { legend: { display: false } } }}
            />
            <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', textAlign: 'center' }}>
              <div style={{ fontSize: '1.5rem', fontWeight: '800' }}>56%</div>
            </div>
          </div>
        </div>
      </div>

      <div className="grid-bottom">
        <div className="chart-panel">
          <div className="chart-title"><Users size={18} color="var(--accent-cyan)" /> Top 5 Risk Entities</div>
          <div style={{ height: '300px' }}>
            <BarChartJS 
              data={{
                labels: entitiesData.map(e => e.name),
                datasets: [{
                  label: 'Risk Count',
                  data: entitiesData.map(e => e.count),
                  backgroundColor: chartColors.severe,
                  borderRadius: 4, barThickness: 24
                }]
              }}
              options={{
                indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } },
                scales: {
                  x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: 'var(--text-muted)' } },
                  y: { grid: { display: false }, ticks: { color: '#fff' } }
                }
              }}
            />
          </div>
        </div>
        <div className="chart-panel">
          <LiveActivityFeed data={data} />
        </div>
      </div>
    </div>
  );
});

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
          <ShieldCheck size={32} color="var(--accent-cyan)" />
          <h2 style={{ margin: 0, fontSize: '1.5rem' }}>Access Control Terminal</h2>
        </div>
        <div style={{ display: 'flex', gap: '10px', marginBottom: '40px' }}>
          <input type="email" placeholder="Enter Corporate Email..." value={newEmail} onChange={(e) => setNewEmail(e.target.value)} style={{ flex: 1, padding: '12px', background: 'var(--panel-bg)', border: '1px solid var(--border-color)', borderRadius: '8px', color: '#fff' }} />
          <button onClick={handleAdd} disabled={loading} className="btn-primary" style={{ background: 'var(--accent-cyan)', color: '#000', border: 'none', padding: '0 20px', borderRadius: '8px', fontWeight: '700' }}>Grant Access</button>
        </div>
        <div className="users-auth-list">
          {authorizedUsers.map(email => (
            <div key={email} className="feed-item" style={{ marginBottom: '10px' }}>
              <span>{email}</span>
              <button onClick={() => handleRemove(email)} style={{ marginLeft: 'auto', background: 'transparent', border: 'none', color: 'var(--danger-red)' }}><Trash2 size={18} /></button>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

function App() {
  const { instance, accounts } = useMsal();
  const [authorizedUsers, setAuthorizedUsers] = useState([]);
  const [authLoading, setAuthLoading] = useState(true);
  const [adminStatus, setAdminStatus] = useState({ isSuperAdmin: false, isAuthorized: false });
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [authError, setAuthError] = useState(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [searchTerm, setSearchTerm] = useState("");
  const API_BASE = import.meta.env.VITE_API_URL || '';

  const metrics = useMemo(() => {
    if (!data || data.length === 0) return { total: 0, riskyUsers: 0, riskPercent: 0, progress: 88.8 };
    const users = [...new Set(data.map(d => d.UserId))];
    const riskyUsers = users.filter(u => data.filter(d => d.UserId === u && (d.Operation?.includes('Delete') || d.Operation?.includes('Share'))).length > 5).length;
    return { total: data.length, riskyUsers, riskPercent: ((riskyUsers / users.length) * 100 || 0).toFixed(1), progress: (Math.min(data.length / 500, 100) * 0.9).toFixed(1) };
  }, [data]);

  const entitiesData = useMemo(() => {
    const userCounts = {};
    data.forEach(d => { userCounts[d.UserId] = (userCounts[d.UserId] || 0) + 1; });
    return Object.entries(userCounts).sort((a,b) => b[1] - a[1]).slice(0, 5).map(([user, count]) => ({ name: user.split('@')[0], count }));
  }, [data]);

  const fetchAuthList = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/admin/authorized-users`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      const list = await response.json();
      const authorizedList = Array.isArray(list) ? list : [];
      setAuthorizedUsers(authorizedList);
      const email = accounts[0]?.username?.toLowerCase();
      const isSuper = ['help-desk@ldplogistics.com', 'kundan@ldplogistics.com'].includes(email);
      setAdminStatus({ isSuperAdmin: isSuper, isAuthorized: authorizedList.includes(email) || isSuper });
    } catch (e) {
      console.error("Auth sync failed", e);
    } finally { setAuthLoading(false); }
  };

  const fetchData = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/audit-logs`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (!response.ok) throw new Error('API Sync Pending');
      setData(await response.json());
      setError(null);
    } catch (err) { setError("Waiting for M365 Backend Sync..."); } finally { setLoading(false); }
  };

  const [webActivity, setWebActivity] = useState([]);
  const [activeCalls, setActiveCalls] = useState([]);
  const [shadowLogs, setShadowLogs] = useState([]);
  const [riskStats, setRiskStats] = useState([]);
  const [profileData, setProfileData] = useState(null);

  const fetchWebActivity = async () => {
    try {
      const res = await fetch(`${API_BASE}/api/security/web-activity`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (res.ok) setWebActivity(await res.json());
    } catch (e) {}
  };

  const fetchActiveCalls = async () => {
    try {
      const res = await fetch(`${API_BASE}/api/security/active-calls`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } });
      if (res.ok) setActiveCalls(await res.json());
    } catch (e) {}
  };

  const fetchSecurityInsights = async () => {
    try {
      const [shadow, risk] = await Promise.all([
        fetch(`${API_BASE}/api/security/shadow-it`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } }),
        fetch(`${API_BASE}/api/security/risk-stats`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } })
      ]);
      if (shadow.ok) setShadowLogs(await shadow.json());
      if (risk.ok) setRiskStats(await risk.json());
    } catch (e) {}
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
    if (accounts.length > 0 && adminStatus.isAuthorized) {
      fetchData();
      if (activeTab === 'web') fetchWebActivity();
      if (activeTab === 'comms') fetchActiveCalls();
      if (activeTab === 'behavior' || activeTab === 'shadow') fetchSecurityInsights();

      const interval = setInterval(() => {
        fetchData();
        if (activeTab === 'web') fetchWebActivity();
        if (activeTab === 'comms') fetchActiveCalls();
        if (activeTab === 'behavior' || activeTab === 'shadow') fetchSecurityInsights();
      }, 5000);
      return () => clearInterval(interval);
    }
  }, [accounts, activeTab, adminStatus.isAuthorized]);

  const handleLogin = () => instance.loginRedirect(loginRequest).catch(e => setAuthError(e.message));

  return (
    <div className="app-layout">
      <AuthenticatedTemplate>
        {authLoading ? (
          <div className="loading-container"><div className="loading-spinner"></div><p>VERIFYING TERMINAL ACCESS...</p></div>
        ) : !adminStatus.isAuthorized ? (
          <div className="access-denied-wrapper fadeIn"><div className="chart-panel"><h1>ACCESS RESTRICTED</h1><button onClick={() => instance.logoutRedirect()} className="btn-primary">Logout</button></div></div>
        ) : (
          <>
            <aside className="sidebar">
              <div className="logo-section"><LayoutDashboard size={24} color="var(--accent-cyan)" /><h2>LDP LOGISTICS</h2></div>
              <div className="sidebar-nav">
                <div style={{ color: 'var(--text-muted)', fontSize: '0.65rem', fontWeight: '800', padding: '0 16px', marginBottom: '8px', opacity: 0.5 }}>MAIN COMMAND</div>
                <button className={`nav-item ${activeTab === 'dashboard' ? 'active' : ''}`} onClick={() => setActiveTab('dashboard')}><LayoutDashboard size={18} /> Overview</button>
                <button className={`nav-item ${activeTab === 'users' ? 'active' : ''}`} onClick={() => setActiveTab('users')}><Users size={18} /> User Intel</button>
                <button className={`nav-item ${activeTab === 'behavior' ? 'active' : ''}`} onClick={() => setActiveTab('behavior')}><ShieldAlert size={18} /> Security Pulse</button>
                
                <div style={{ color: 'var(--text-muted)', fontSize: '0.65rem', fontWeight: '800', padding: '0 16px', marginTop: '20px', marginBottom: '8px', opacity: 0.5 }}>MONITORING</div>
                <button className={`nav-item ${activeTab === 'comms' ? 'active' : ''}`} onClick={() => setActiveTab('comms')}><Activity size={18} /> Comms Intel</button>
                <button className={`nav-item ${activeTab === 'web' ? 'active' : ''}`} onClick={() => setActiveTab('web')}><Globe size={18} /> Web Activity</button>
                <button className={`nav-item ${activeTab === 'shadow' ? 'active' : ''}`} onClick={() => setActiveTab('shadow')}><FileText size={18} /> Shadow IT</button>

                {adminStatus.isSuperAdmin && (
                  <button className={`nav-item ${activeTab === 'admin' ? 'active' : ''}`} onClick={() => setActiveTab('admin')} style={{ marginTop: 'auto', color: 'var(--warning-yellow)' }}><Settings size={18} /> Admin Terminal</button>
                )}
              </div>
            </aside>

            <main className="main-content">
              <div className="top-header">
                <div className="system-status">
                  <span className="pulse-dot green"></span>
                  <span style={{ color: 'var(--success-green)', fontWeight: '800', fontSize: '0.7rem' }}>SECURE CONNECTION ACTIVE</span>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                  <div className="user-pill"><UserIcon size={14} /><span>{accounts[0]?.name?.split(' ')[0]}</span></div>
                  <SystemClock />
                  <button onClick={() => instance.logoutRedirect()} className="logout-icon-btn"><LogIn size={20} /></button>
                </div>
              </div>

              <div className="dashboard-content">
                {activeTab === 'dashboard' && <DashboardContent data={data} metrics={metrics} entitiesData={entitiesData} handleUserClick={(e) => setActiveTab('users')} />}
                
                {activeTab === 'users' && (
                  <div className="fadeIn">
                    <div className="chart-panel">
                      <div className="chart-title">Employee Intelligence Overview</div>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th>Identity</th>
                            <th>Status</th>
                            <th>Risk Level</th>
                            <th>Activity Count</th>
                            <th>Actions</th>
                          </tr>
                        </thead>
                        <tbody>
                          {userIntelligenceSummary.map(u => (
                            <tr key={u.id}>
                              <td><div style={{ fontWeight: 700 }}>{u.id.split('@')[0]}</div><div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{u.id}</div></td>
                              <td><span className={`pill ${u.isLive ? 'pill-cyan' : 'pill-gray'}`}>{u.isLive ? 'LIVE' : 'IDLE'}</span></td>
                              <td><span className={`pill ${u.actions > 50 ? 'pill-red' : 'pill-cyan'}`}>{u.actions > 50 ? 'CRITICAL' : 'LOW'}</span></td>
                              <td>{u.actions} hits</td>
                              <td><button className="btn-primary" style={{ padding: '4px 12px', borderRadius: '4px', background: 'rgba(0,212,255,0.1)', color: 'var(--accent-cyan)', border: '1px solid var(--accent-cyan)' }}>Investigate</button></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {activeTab === 'admin' && <AdminPortal authorizedUsers={authorizedUsers} fetchAuthList={fetchAuthList} API_BASE={API_BASE} />}
                
                {/* Fallback for other tabs being modernized */}
                {['behavior', 'comms', 'web', 'shadow'].includes(activeTab) && (
                  <div className="fadeIn" style={{ textAlign: 'center', padding: '100px' }}>
                    <Activity size={48} color="var(--accent-cyan)" />
                    <h2 style={{ marginTop: '20px' }}>{activeTab.toUpperCase()} TERMINAL ACTIVE</h2>
                    <p style={{ color: 'var(--text-muted)' }}>Real-time telemetry and forensic logs are being streamed to this terminal.</p>
                  </div>
                )}
              </div>
            </main>
          </>
        )}
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <div className="login-screen">
          <div className="login-container">
            <div className="chart-panel" style={{ padding: '40px', textAlign: 'center' }}>
              <ShieldAlert size={48} color="var(--accent-cyan)" style={{ marginBottom: '20px' }} />
              <h1>SECURITY TERMINAL</h1>
              <p>LDP LOGISTICS DATA GOVERNANCE</p>
              <button onClick={handleLogin} className="btn-primary" style={{ width: '100%', marginTop: '20px' }}>Login with Microsoft</button>
              {authError && <div style={{ color: 'var(--danger-red)', marginTop: '10px' }}>{authError}</div>}
            </div>
          </div>
        </div>
      </UnauthenticatedTemplate>
    </div>
  );
}

export default App;
