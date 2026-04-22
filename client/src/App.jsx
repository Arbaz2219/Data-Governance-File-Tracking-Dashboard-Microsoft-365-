// v1.0.2 - Forcing clean build to resolve potential cache issues
import React, { useState, useEffect, useMemo } from 'react';
import { Activity, Users, FileText, Share2, LogIn, LayoutDashboard, Database, User as UserIcon, ShieldAlert, Laptop, Globe, Lock } from 'lucide-react';
import { ResponsiveSankey } from '@nivo/sankey';
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, BarChart, Bar } from 'recharts';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import './App.css';

const SystemClock = () => {
    const [time, setTime] = useState(new Date());
    useEffect(() => {
        const timer = setInterval(() => setTime(new Date()), 1000);
        return () => clearInterval(timer);
    }, []);

    const dateStr = time.toLocaleDateString('en-US', { year: 'numeric', month: 'numeric', day: 'numeric' });
    const timeStr = time.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true }).toLowerCase();

    return (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', color: '#fff', padding: '0 15px', opacity: 0.9 }}>
            <span style={{ fontSize: '1rem', fontWeight: '600', letterSpacing: '0.5px' }}>{timeStr}</span>
            <span style={{ fontSize: '0.6rem', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{dateStr}</span>
        </div>
    );
};

const Skeleton = ({ width, height, borderRadius = "8px" }) => (
  <div className="skeleton" style={{ width, height, borderRadius, marginBottom: '10px' }} />
);

const SecurityCharts = React.memo(({ sankeyData, activityData, error }) => {
  return (
    <div className="charts-grid">
      <div className="glass-panel chart-panel">
        <h2>Data Flow Architecture (User → Action → Platform)</h2>
        <div className="chart-container" style={{ height: "320px", width: "100%" }}>
          {error ? <div style={{color: 'red', marginTop: '20px'}}>{error}</div> : (
            <ResponsiveSankey
                data={sankeyData}
                margin={{ top: 20, right: 180, bottom: 20, left: 180 }}
                align="justify"
                colors={{ scheme: 'set2' }}
                nodeOpacity={1}
                nodeThickness={20}
                nodeSpacing={20}
                nodeBorderWidth={0}
                linkOpacity={0.3}
                linkHoverOpacity={0.6}
                labelPadding={15}
                labelTextColor="var(--text-main)"
                theme={{
                  labels: { text: { fontSize: 11, fontWeight: 700, fontFamily: 'Inter'} },
                  tooltip: { container: { background: "#1e293b", color: "#f8fafc", borderRadius: '8px', border: "1px solid rgba(255,255,255,0.1)", boxShadow: '0 10px 15px -3px rgba(0,0,0,0.4)' } }
                }}
            />
          )}
        </div>
      </div>

      <div className="glass-panel chart-panel">
        <h2>Security Activity Pulse</h2>
        <div className="chart-container" style={{ height: "400px", width: "100%" }}>
           <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={activityData} margin={{ top: 20, right: 20, bottom: 20, left: -20 }}>
              <defs>
                <linearGradient id="colorEvents" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="var(--primary)" stopOpacity={0.5}/>
                  <stop offset="95%" stopColor="var(--primary)" stopOpacity={0}/>
                </linearGradient>
              </defs>
              <XAxis dataKey="hour" stroke="#8892b0" fontSize={11} tickLine={false} axisLine={false} />
              <YAxis stroke="#8892b0" fontSize={11} tickLine={false} axisLine={false} />
              <Tooltip 
                contentStyle={{background: 'rgba(31, 34, 64, 0.9)', border: '1px solid var(--primary)', borderRadius: '8px', fontSize: '0.85rem'}}
                itemStyle={{color: 'var(--primary)'}}
              />
              <Area type="monotone" dataKey="events" stroke="var(--primary)" strokeWidth={3} fillOpacity={1} fill="url(#colorEvents)" />
              </AreaChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
});

const CyberGauge = ({ value }) => {
  const rotation = (value / 100) * 180 - 90;
  return (
    <div style={{ position: 'relative', width: '200px', height: '120px', margin: 'auto' }}>
      <svg width="200" height="120" viewBox="0 0 200 120">
        <path d="M 20 100 A 80 80 0 0 1 180 100" fill="none" stroke="#222" strokeWidth="20" />
        <path d="M 20 100 A 80 80 0 0 1 73 35" fill="none" stroke="#52c41a" strokeWidth="20" />
        <path d="M 73 35 A 80 80 0 0 1 127 35" fill="none" stroke="#faad14" strokeWidth="20" />
        <path d="M 127 35 A 80 80 0 0 1 180 100" fill="none" stroke="#f5222d" strokeWidth="20" />
        <line x1="100" y1="100" x2={100 + 70 * Math.sin((rotation * Math.PI) / 180)} y2={100 - 70 * Math.cos((rotation * Math.PI) / 180)} stroke="#fff" strokeWidth="4" strokeLinecap="round" />
        <circle cx="100" cy="100" r="8" fill="#fff" />
      </svg>
      <div style={{ textAlign: 'center', marginTop: '-20px', fontWeight: '800', fontSize: '1.2rem' }}>{value}%</div>
    </div>
  );
};

const CyberMetricCard = ({ title, value }) => (
  <div className="metric-card">
    <div className="card-header-bar">{title}</div>
    <div className="card-content">
      <div style={{ fontSize: '2.5rem', fontWeight: '800', letterSpacing: '1px' }}>{value}</div>
    </div>
  </div>
);

const KpiCard = ({ label, value }) => (
  <div className="metric-card headerless">
    <div className="label">{label}</div>
    <div className="value">{value}</div>
  </div>
);

const HeatMapTable = ({ data }) => {
  const rows = ['Severe', 'Major', 'Moderate', 'Minor', 'Insignificance'];
  const cols = ['Rare', 'Unlikely', 'Moderate', 'Likely', 'Almost certain'];
  
  // Dynamic cell generation for "Heart Map"
  const getCellClass = (row, colIndex) => {
    if (row === 'Severe' || colIndex === 4) return 'cell-severe';
    if (row === 'Major' || colIndex === 3) return 'cell-major';
    if (row === 'Moderate' || colIndex === 2) return 'cell-moderate';
    if (row === 'Minor' || colIndex === 1) return 'cell-minor';
    return 'cell-insignificant';
  };

  return (
    <table className="heart-map-table">
      <thead>
        <tr>
          <th className="header-label">Total # of Risk Rating</th>
          {cols.map(c => <th key={c} className="header-label">{c}</th>)}
        </tr>
      </thead>
      <tbody>
        {rows.map((r, i) => (
          <tr key={r}>
            <td className="header-label">{r}</td>
            {cols.map((c, j) => (
              <td key={c} className={getCellClass(r, j)}>
                {Math.floor(Math.random() * 100) + 20}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
};

const DashboardContent = React.memo(({ data, error, handleUserClick, sankeyData, activityData, searchTerm }) => {
  const metrics = useMemo(() => {
    if (!data || data.length === 0) return { total: 0, riskyUsers: 0, riskPercent: 0, progress: 88.8 };
    const users = [...new Set(data.map(d => d.UserId))];
    const riskyUsers = users.filter(u => data.filter(d => d.UserId === u && (d.Operation?.includes('Delete') || d.Operation?.includes('Share'))).length > 5).length;
    return { 
      total: data.length, 
      riskyUsers, 
      riskPercent: ((riskyUsers / users.length) * 100 || 0).toFixed(1),
      progress: (Math.min(data.length / 500, 100) * 0.9).toFixed(1)
    };
  }, [data]);

  const riskBreakdown = [
    { name: 'Critical', value: 14, fill: '#ff4d4f' },
    { name: 'High', value: 33, fill: '#faad14' },
    { name: 'Medium', value: 48, fill: '#00f2ff' },
    { name: 'Low', value: 5, fill: '#52c41a' },
  ];

  const actionPlanData = [
    { name: 'Implemented', value: 30.9, fill: '#00f2ff' },
    { name: 'Planned', value: 8.5, fill: 'rgba(0, 242, 255, 0.4)' },
    { name: 'Tbd', value: 50.2, fill: 'rgba(255, 255, 255, 0.1)' },
  ];

  const vulnsData = useMemo(() => {
    const counts = { 'Encryption vulns': 26, 'Excessive permissions': 68, 'Dormant accounts': 34, 'Unauthorized Apps': 45, 'Vpn Usage': 29 };
    return Object.entries(counts).map(([name, count]) => ({ name, count }));
  }, []);

  const entitiesData = useMemo(() => {
    const userCounts = {};
    data.forEach(d => { userCounts[d.UserId] = (userCounts[d.UserId] || 0) + 1; });
    return Object.entries(userCounts)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5)
      .map(([user, count]) => ({ name: user.split('@')[0], count }));
  }, [data]);

  const formatNJTime = (creationTime) => {
    if (!creationTime) return "N/A";
    const date = new Date(creationTime);
    const dateStr = date.toLocaleDateString('en-US', { year: 'numeric', month: 'numeric', day: 'numeric' });
    const timeStr = date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true }).toLowerCase();
    return `${dateStr} time ${timeStr}`;
  };

  const liveFeedData = useMemo(() => {
    if (!data || data.length === 0) return [];
    let processed = [...data].sort((a,b) => new Date(b.CreationTime).getTime() - new Date(a.CreationTime).getTime());
    
    if (searchTerm) {
        const term = searchTerm.toLowerCase();
        processed = processed.filter(d => 
            d.UserId?.toLowerCase().includes(term) || 
            d.ObjectId?.toLowerCase().includes(term) || 
            d.Operation?.toLowerCase().includes(term)
        );
    }

    return processed.filter(d => !d.UserId?.includes('app@sharepoint')).slice(0, 50).map(event => {
      let filename = event.ObjectId?.split('/').pop();
      try { filename = decodeURIComponent(filename); } catch (e) {}
      if (!filename || filename.includes('http')) filename = 'Folder / System Item';
      
      return {
        id: event.Id,
        rawEmail: event.UserId,
        time: formatNJTime(event.CreationTime),
        user: event.UserId?.split('@')[0] || 'System',
        action: event.Operation || 'Unknown',
        target: filename.substring(0, 50),
        location: event.Workload || 'SharePoint'
      };
    });
  }, [data]);

  return (
    <div style={{ animation: 'fadeIn 0.5s ease-out' }}>
      {/* ROW 1: KPIs */}
      <div className="kpi-row">
        <KpiCard label="% Risks >= threshold" value={`${metrics.riskPercent}%`} />
        <KpiCard label="# Of risks >= threshold" value={metrics.riskyUsers} />
        <KpiCard label="Risks analysis progress" value={`${metrics.progress}%`} />
        <KpiCard label="Response progress" value="52.6%" />
      </div>

      {/* ROW 2: Analysis Tiers */}
      <div className="analysis-row">
        <div className="glass-panel chart-panel">
          <h2 style={{ background: '#00f2ff', color: '#000' }}>Risk rating breakdown</h2>
          <div style={{ flex: 1, padding: '20px' }}>
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={riskBreakdown} innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                  {riskBreakdown.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.fill} />)}
                </Pie>
                <Tooltip contentStyle={{ background: '#001529', border: 'none' }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="glass-panel chart-panel">
          <h2 style={{ background: '#00f2ff', color: '#000' }}>Risk heart map</h2>
          <div style={{ flex: 1, padding: '15px' }}>
            <HeatMapTable />
          </div>
        </div>

        <div className="glass-panel chart-panel">
          <h2 style={{ background: '#00f2ff', color: '#000' }}>Action plan breakdown</h2>
          <div style={{ flex: 1, padding: '20px' }}>
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={actionPlanData} innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                  {actionPlanData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.fill} />)}
                </Pie>
                <Tooltip contentStyle={{ background: '#001529', border: 'none' }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* ROW 3: Targets */}
      <div className="targets-row">
        <div className="glass-panel chart-panel">
          <h2 style={{ background: '#00f2ff', color: '#000' }}>#Risks &gt;= threshold: top 5 vulnerabilities</h2>
          <div style={{ flex: 1, padding: '20px' }}>
             <ResponsiveContainer width="100%" height={200}>
              <BarChart layout="vertical" data={vulnsData}>
                <XAxis type="number" hide />
                <YAxis dataKey="name" type="category" stroke="#fff" fontSize={10} width={120} />
                <Tooltip contentStyle={{ background: '#001529', border: '1px solid #ff4d4f' }} />
                <Bar dataKey="count" fill="#ff4d4f" radius={[0, 4, 4, 0]} barSize={20} />
              </BarChart>
             </ResponsiveContainer>
          </div>
        </div>

        <div className="glass-panel chart-panel">
          <h2 style={{ background: '#00f2ff', color: '#000' }}># Risks &gt;= threshold: top 5 entities</h2>
          <div style={{ flex: 1, padding: '20px' }}>
             <ResponsiveContainer width="100%" height={200}>
              <BarChart layout="vertical" data={entitiesData}>
                <XAxis type="number" hide />
                <YAxis dataKey="name" type="category" stroke="#fff" fontSize={10} width={120} />
                <Tooltip contentStyle={{ background: '#001529', border: '1px solid #00f2ff' }} />
                <Bar dataKey="count" fill="#00f2ff" radius={[0, 4, 4, 0]} barSize={20} />
              </BarChart>
             </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="glass-panel chart-panel" style={{ marginTop: '2rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 20px', background: '#000c17', borderBottom: '1px solid rgba(0,242,255,0.2)' }}>
          <h2 style={{ background: 'none', padding: 0, textAlign: 'left', fontSize: '1rem', display: 'flex', alignItems: 'center', gap: '10px', color: '#00f2ff' }}>
            <span className="status-pulse success" style={{ background: '#00f2ff' }}></span>
            Real-time Cyber Risk Matrix
          </h2>
        </div>
        <div className="table-container" style={{ padding: '0' }}>
          <table className="custom-table" style={{ width: '100%' }}>
            <thead style={{ background: '#001a33' }}>
              <tr style={{ textAlign: 'left', color: '#8c8c8c', fontSize: '0.8rem', textTransform: 'uppercase' }}>
                <th style={{ padding: '15px' }}>DateTime</th>
                <th>Actor Entity</th>
                <th>Risk Classification</th>
                <th>Target Resource</th>
              </tr>
            </thead>
            <tbody>
              {liveFeedData.slice(0, 15).map((row) => (
                <tr key={row.id}>
                  <td style={{ color: 'var(--text-muted)', fontSize: '0.85rem', padding: '12px 15px' }}>{row.time}</td>
                  <td onClick={() => handleUserClick(row.rawEmail)} style={{ color: '#00f2ff', fontWeight: '600', cursor: 'pointer' }}>{row.user}</td>
                  <td>
                    <span style={{ 
                      padding: '4px 10px', 
                      borderRadius: '4px', 
                      fontSize: '0.7rem',
                      fontWeight: '700',
                      background: row.action.includes('Delete') ? 'rgba(245, 34, 45, 0.1)' : 'rgba(0, 242, 255, 0.1)',
                      color: row.action.includes('Delete') ? '#f5222d' : '#00f2ff',
                      border: `1px solid ${row.action.includes('Delete') ? 'rgba(245, 34, 45, 0.2)' : 'rgba(0, 242, 255, 0.2)'}`
                    }}>
                      {row.action}
                    </span>
                  </td>
                  <td style={{ fontSize: '0.8rem', opacity: 0.8 }}>{row.target}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
});

function App() {
  const { instance, accounts, inProgress } = useMsal();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [authError, setAuthError] = useState(null);
  const [selectedUser, setSelectedUser] = useState(null);
  const [fullProfile, setFullProfile] = useState(null);
  const [webActivity, setWebActivity] = useState([]);
  const [activeTab, setActiveTab] = useState('feed'); 
  const [searchTerm, setSearchTerm] = useState("");
  const [usersList, setUsersList] = useState([]);
  const [selectedUserIntel, setSelectedUserIntel] = useState(null);
  const [activeCalls, setActiveCalls] = useState([]);
  const [shadowLogs, setShadowLogs] = useState([]);
  const [riskStats, setRiskStats] = useState([]);

  const isAuthorized = accounts[0]?.username?.toLowerCase().endsWith('@ldplogistics.com');

  useEffect(() => {
    if (accounts.length > 0) {
      setAuthError(null);
    }
  }, [accounts]);

  const API_BASE = import.meta.env.VITE_API_URL || '';
  const fetchData = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/audit-logs`, {
        headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' }
      });
      if (!response.ok) throw new Error('API Sync Pending');
      const json = await response.json();
      setData(json);
      setError(null);
    } catch (err) {
      setError("Waiting for M365 Backend Sync...");
    } finally {
      setLoading(false);
    }
  };

  const sankeyData = useMemo(() => {
    if (!data || data.length === 0) return { nodes: [{id: 'Loading'}, {id: 'Data'}], links: [{source: 'Loading', target: 'Data', value: 1}] };
    const nodesMap = new Map();
    const addNode = (id) => { if (id && !nodesMap.has(id)) nodesMap.set(id, { id }) };
    const linksMap = new Map();
    const addLink = (source, target) => {
        if (!source || !target || source === target) return;
        addNode(source);
        addNode(target);
        const key = `${source}->${target}`;
        linksMap.set(key, { source, target, value: (linksMap.get(key)?.value || 0) + 1 });
    };
    const topData = [...data].reverse().filter(d => !d.UserId?.includes('app@sharepoint') && !d.UserId?.includes('urn:spo')).slice(0, 40);
    topData.forEach(event => {
      let user = event.UserId?.split('@')[0] || "Unknown";
      if(user.length > 25) user = user.substring(0, 25) + '...';
      let action = event.Operation?.replace('File', '') || "Other";
      const location = event.Workload === 'OneDrive' ? 'OneDrive Storage' : 'SharePoint Site';
      action = `Action: ${action}`;
      addLink(user, action);
      addLink(action, location);
    });
    return { nodes: Array.from(nodesMap.values()), links: Array.from(linksMap.values()) };
  }, [data]);

  const activityData = useMemo(() => {
    if (!data || data.length === 0) return [];
    const counts = {};
    data.forEach(d => {
      if(!d.CreationTime) return;
      const dateInEst = new Date(d.CreationTime).toLocaleString('en-US', { timeZone: 'America/New_York', hour: 'numeric', hour12: false });
      const hour = parseInt(dateInEst);
      counts[hour] = (counts[hour] || 0) + 1;
    });
    return Object.keys(counts).map(h => ({ hour: `${h}:00`, events: counts[h] })).sort((a,b) => parseInt(a.hour) - parseInt(b.hour));
  }, [data]);

  const fetchWebActivity = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/security/web-activity`, {
        headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' }
      });
      if (response.ok) {
        const json = await response.json();
        setWebActivity(json);
      }
    } catch (err) { console.error("Web fetch error:", err); }
  };

  const filteredWebActivity = useMemo(() => {
    if (!webActivity || webActivity.length === 0) return [];
    if (!searchTerm) return webActivity;
    const term = searchTerm.toLowerCase();
    return webActivity.filter(log => 
      log.DeviceName?.toLowerCase().includes(term) || 
      log.RemoteUrl?.toLowerCase().includes(term) || 
      log.InitiatingProcessAccountName?.toLowerCase().includes(term)
    );
  }, [webActivity, searchTerm]);

  useEffect(() => {
    if (accounts.length > 0) {
      fetchData();
      fetchWebActivity();
      const interval = setInterval(() => {
        fetchData();
        if (activeTab === 'web') fetchWebActivity();
      }, 3000); 
      return () => clearInterval(interval);
    }
  }, [accounts]);

  const fetchActiveCalls = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/security/active-calls`, {
        headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' }
      });
      if (response.ok) {
        const json = await response.json();
        setActiveCalls(json);
      }
    } catch (err) { console.error("Call fetch error:", err); }
  };

  const fetchSecurityInsights = async () => {
    try {
      const [shadowRes, riskRes] = await Promise.all([
        fetch(`${API_BASE}/api/security/shadow-it`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } }),
        fetch(`${API_BASE}/api/security/risk-stats`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } })
      ]);
      if (shadowRes.ok) setShadowLogs(await shadowRes.json());
      if (riskRes.ok) setRiskStats(await riskRes.json());
    } catch (err) { console.error("Insights fetch error:", err); }
  };

  useEffect(() => {
    if (accounts.length > 0) {
      fetchData();
      fetchWebActivity();
      fetchActiveCalls();
      fetchSecurityInsights();
      const interval = setInterval(() => {
        fetchData();
        if (activeTab === 'web') fetchWebActivity();
        if (activeTab === 'comms') fetchActiveCalls();
        if (activeTab === 'shadow' || activeTab === 'behavior') fetchSecurityInsights();
      }, 5000); 
      return () => clearInterval(interval);
    }
  }, [accounts, activeTab]);

  useEffect(() => {
    if (accounts.length > 0) {
      fetch(`${API_BASE}/api/users/list`, { headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' } })
        .then(res => res.json())
        .then(data => setUsersList(data))
        .catch(err => console.error("Failed to fetch users list", err));
    }
  }, [accounts]);

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch(e => {
      console.error("[MSAL Login Error]", e);
      setAuthError(e.errorMessage || e.message || "Login failed. Please try again.");
    });
  };

  const handleUserClick = async (email) => {
    if (!email || email === 'System') return;
    let fullEmail = email.includes('@') ? email : `${email}@ldplogistics.com`;
    setSelectedUser(fullEmail);
    setFullProfile(null);
    try {
        const profileResponse = await fetch(`${API_BASE}/api/user/${fullEmail}/profile`, {
          headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' }
        });
        if(!profileResponse.ok) throw new Error();
        const profileData = await profileResponse.json();
        setFullProfile(profileData);
    } catch(e) {
        setFullProfile({ error: true });
    }
  };

  const handleReboot = async (deviceId, deviceName) => {
    if (!window.confirm(`🚨 CRITICAL ACTION: Are you sure you want to REMOTE RESTART the device "${deviceName}"? \n\nAny unsaved work for this user will be LOST immediately.`)) return;
    try {
        const res = await fetch(`${API_BASE}/api/device/${deviceId}/reboot`, { 
            method: 'POST',
            headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' }
        });
        const resData = await res.json();
        if (resData.success) {
            alert(`✅ Command Dispatched: ${deviceName} is being restarted.`);
        } else {
            throw new Error(resData.error);
        }
    } catch (e) {
        alert("❌ Failed to trigger restart: " + (e.message || "Permissions Denied"));
    }
  };

  const formatNJTime = (creationTime) => {
    if (!creationTime) return "N/A";
    const date = new Date(creationTime);
    const dateStr = date.toLocaleDateString('en-US', { year: 'numeric', month: 'numeric', day: 'numeric' });
    const timeStr = date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true }).toLowerCase();
    return `${dateStr} time ${timeStr}`;
  };

  return (
    <div className="app-layout">
      <AuthenticatedTemplate>
        <aside className="sidebar">
          <div className="logo-section">
            <div style={{ padding: '8px', background: 'rgba(20, 184, 166, 0.1)', borderRadius: '10px', display: 'flex' }}>
              <LayoutDashboard size={24} color="var(--primary)" />
            </div>
            <h2>LDP LOGISTICS</h2>
          </div>
          
          <nav className="nav-menu">
            <div className="nav-label" style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: '700', textTransform: 'uppercase', marginBottom: '0.75rem', paddingLeft: '1.5rem', letterSpacing: '1.5px' }}>Terminal Menu</div>
            <div 
              className={`nav-item ${activeTab === 'feed' ? 'active' : ''}`} 
              onClick={() => setActiveTab('feed')}
            >
              <LayoutDashboard size={20} /> Dashboard
            </div>
            <div 
              className={`nav-item ${activeTab === 'users' ? 'active' : ''}`} 
              onClick={() => setActiveTab('users')}
            >
              <Users size={20} /> User Intelligence
            </div>
            <div 
              className={`nav-item ${activeTab === 'behavior' ? 'active' : ''}`} 
              onClick={() => setActiveTab('behavior')}
            >
              <Activity size={20} /> User Behavior
            </div>

            
          </nav>

        </aside>

        <main className="main-content" style={{ padding: 0 }}>
          <header className="topbar">
            <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
               <span style={{ fontWeight: '600', fontSize: '1rem' }}>Data Governance Dashboard</span>
               <div className="search-bar-wrapper" style={{ position: 'relative', width: '400px' }}>
                  <input 
                     type="text" 
                     placeholder="Search resources, users, and logs (G+/)" 
                     value={searchTerm}
                     onChange={(e) => setSearchTerm(e.target.value)}
                  />
               </div>
            </div>
            
            <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
              <SystemClock />
              <div className="top-icons" style={{ display: 'flex', gap: '15px', color: '#fff', opacity: 0.8 }}>
                <Activity size={18} style={{ cursor: 'pointer' }} />
                <FileText size={18} style={{ cursor: 'pointer' }} />
                <Globe size={18} style={{ cursor: 'pointer' }} />
              </div>
              
              <div className="profile-pill" style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'rgba(255,255,255,0.1)', padding: '4px 12px', borderRadius: '4px' }}>
                <div style={{ width: '24px', height: '24px', background: '#004578', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  <UserIcon size={14} color="#fff" />
                </div>
                <span style={{ fontSize: '0.8rem', fontWeight: '600' }}>{accounts[0]?.name?.split(' ')[0]}</span>
              </div>
              
              <button 
                onClick={() => instance.logoutRedirect()}
                style={{ background: 'rgba(255, 255, 255, 0.1)', border: 'none', color: '#fff', padding: '6px 12px', borderRadius: '2px', fontWeight: '600', cursor: 'pointer', fontSize: '0.75rem' }}
              >
                Sign out
              </button>
            </div>
          </header>

          {!isAuthorized ? (
            <div className="glass-panel" style={{ padding: '5rem', textAlign: 'center' }}>
               <ShieldAlert size={64} color="var(--error)" style={{ marginBottom: '1.5rem' }} />
               <h2 style={{ color: 'var(--text-main)', fontSize: '2rem' }}>Access Denied</h2>
               <p style={{ color: 'var(--text-muted)', margin: '1rem 0' }}>This terminal is restricted to authorized Security Personnel only.</p>
               <p style={{ color: 'var(--error)', fontWeight: 'bold' }}>Contact the Head of IT to request access for {accounts[0]?.username}.</p>
            </div>
          ) : loading ? (
            <div style={{ padding: '20px' }}>
              <div className="metrics-grid">
                <Skeleton width="100%" height="120px" borderRadius="16px" />
                <Skeleton width="100%" height="120px" borderRadius="16px" />
                <Skeleton width="100%" height="120px" borderRadius="16px" />
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem', marginTop: '20px' }}>
                <Skeleton width="100%" height="400px" borderRadius="16px" />
                <Skeleton width="100%" height="400px" borderRadius="16px" />
              </div>
            </div>
          ) : (
            <>
              {activeTab === 'feed' && (
                <DashboardContent 
                  data={data} 
                  error={error} 
                  handleUserClick={handleUserClick} 
                  sankeyData={sankeyData}
                  activityData={activityData}
                  searchTerm={searchTerm}
                />
              )}


  



            {activeTab === 'users' && (
              <div className="user-intelligence-container" style={{ display: 'grid', gridTemplateColumns: '300px 1fr', gap: '2rem', height: 'calc(100vh - 160px)' }}>
                {/* User List Sidebar */}
                <div className="glass-panel" style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                   <div style={{ padding: '20px' }}>
                      <h3 style={{ margin: '0 0 15px 0' }}>Employee Directory</h3>
                      <input 
                        type="text" 
                        placeholder="Search users..." 
                        style={{ width: '100%', padding: '10px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--glass-border)', borderRadius: '8px', color: '#fff' }}
                        onChange={(e) => setSearchTerm(e.target.value)}
                      />
                   </div>
                   <div style={{ flex: 1, overflowY: 'auto', padding: '0 10px 20px' }}>
                      {usersList.filter(u => u.toLowerCase().includes(searchTerm.toLowerCase())).map(user => (
                        <div 
                          key={user}
                          className={`nav-item ${selectedUserIntel === user ? 'active' : ''}`}
                          style={{ marginBottom: '5px', padding: '12px' }}
                          onClick={() => {
                            setSelectedUserIntel(user);
                            handleUserClick(user); // Also trigger the existing profile fetch
                          }}
                        >
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                             <div style={{ width: '32px', height: '32px', background: 'var(--primary)', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px' }}>
                                {user[0].toUpperCase()}
                             </div>
                             <div style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                <div style={{ fontWeight: '600', fontSize: '0.9rem' }}>{user.split('@')[0]}</div>
                                <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{user}</div>
                             </div>
                          </div>
                        </div>
                      ))}
                   </div>
                </div>

                {/* User Detail Content */}
                <div style={{ height: '100%', overflowY: 'auto' }}>
                   {selectedUserIntel ? (
                      <div style={{ animation: 'fadeIn 0.5s ease' }}>
                         <div className="glass-panel" style={{ padding: '30px', marginBottom: '2rem' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                               <div>
                                  <h1 style={{ margin: 0 }}>{fullProfile?.profile?.displayName || selectedUserIntel.split('@')[0]}</h1>
                                  <p style={{ color: 'var(--primary)', fontWeight: '600' }}>{selectedUserIntel}</p>
                                  <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
                                     <span className="glass-panel" style={{ padding: '4px 12px', fontSize: '0.8rem', background: 'rgba(130, 87, 229, 0.1)' }}>{fullProfile?.profile?.jobTitle || 'Standard User'}</span>
                                     <span className="glass-panel" style={{ padding: '4px 12px', fontSize: '0.8rem', background: 'rgba(0, 210, 255, 0.1)' }}>{fullProfile?.profile?.department || 'Operations'}</span>
                                  </div>
                               </div>
                               <div style={{ textAlign: 'right' }}>
                                  <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>M365 STATUS</div>
                                  <div style={{ color: 'var(--success)', fontWeight: '800' }}>ACTIVE & MONITORED</div>
                               </div>
                            </div>
                         </div>

                         <div className="metrics-grid">
                            <div className="glass-panel" style={{ padding: '25px', display: 'flex', gap: '20px', alignItems: 'center' }}>
                               <div style={{ width: '60px', height: '60px', background: 'linear-gradient(45deg, var(--primary), var(--secondary))', borderRadius: '16px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff' }}>
                                  <Database size={32} />
                               </div>
                               <div>
                                  <div style={{ fontSize: '2rem', fontWeight: '800' }}>{fullProfile?.stats?.uniqueFileCount || '0'}</div>
                                  <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>OneDrive Files Accessible</div>
                               </div>
                            </div>
                            <div className="glass-panel" style={{ padding: '25px', display: 'flex', gap: '20px', alignItems: 'center' }}>
                               <div style={{ width: '60px', height: '60px', background: 'rgba(255, 51, 102, 0.1)', borderRadius: '16px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--error)' }}>
                                  <ShieldAlert size={32} />
                               </div>
                               <div>
                                  <div style={{ fontSize: '2rem', fontWeight: '800' }}>{fullProfile?.stats?.totalEvents || '0'}</div>
                                  <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>Total Audit Transactions</div>
                               </div>
                            </div>
                         </div>

                         {fullProfile && !fullProfile.error && (
                            <div className="glass-panel" style={{ padding: '25px', marginTop: '2rem' }}>
                               <h3>Shared Resources Access Summary</h3>
                               <div className="storage-box" style={{ background: 'transparent', padding: 0 }}>
                                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}>
                                     <span>OneDrive Storage Utilization</span>
                                     <span>{(fullProfile.storage.used / 1073741824).toFixed(2)} GB / {(fullProfile.storage.total / 1073741824).toFixed(0)} GB</span>
                                  </div>
                                  <div className="storage-bar-bg">
                                     <div className="storage-bar-fill" style={{ width: `${(fullProfile.storage.used / fullProfile.storage.total * 100)}%` }}></div>
                                  </div>
                               </div>
                            </div>
                         )}

                         <div className="glass-panel" style={{ marginTop: '2rem', padding: '25px' }}>
                            <h3>Recent Security Pulse (Last 50 Actions)</h3>
                            <div className="table-container">
                               <table className="custom-table" style={{ width: '100%' }}>
                                  <thead>
                                     <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.8rem' }}>
                                        <th>DateTime</th>
                                        <th>Action</th>
                                        <th>Resource Path</th>
                                     </tr>
                                  </thead>
                                  <tbody>
                                     {data.filter(d => d.UserId?.toLowerCase() === selectedUserIntel.toLowerCase()).slice(0, 10).map((log, i) => (
                                        <tr key={i}>
                                           <td>{new Date(log.CreationTime).toLocaleString()}</td>
                                           <td><span style={{ color: 'var(--primary)' }}>{log.Operation}</span></td>
                                           <td style={{ fontSize: '0.8rem', opacity: 0.8 }}>{log.ObjectId?.split('/').pop()}</td>
                                        </tr>
                                     ))}
                                  </tbody>
                               </table>
                            </div>
                         </div>
                      </div>
                   ) : (
                      <div className="glass-panel" style={{ height: '100%', display: 'flex', flexWrap: 'wrap', placeContent: 'center', textAlign: 'center', opacity: 0.5 }}>
                         <div>
                            <Users size={64} style={{ marginBottom: '1rem' }} />
                            <h2>Select a user to analyze intelligence</h2>
                            <p>Real-time file access and OneDrive permissions will be calculated upon selection.</p>
                         </div>
                      </div>
                   )}
                </div>
              </div>
            )}

            {activeTab === 'behavior' && (
              <div className="behavior-container" style={{ animation: 'fadeIn 0.5s ease-out', padding: '20px' }}>
                <div className="glass-panel" style={{ padding: '25px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                     <h2 style={{ margin: 0, color: 'var(--primary)', display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <Activity size={24} /> User Behavior & Risk Analysis
                     </h2>
                     <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                        Aggregated from last 24h M365 Audit Logs
                     </div>
                  </div>

                  <div className="table-container">
                    <table className="custom-table" style={{ width: '100%' }}>
                      <thead>
                        <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.8rem', textTransform: 'uppercase' }}>
                          <th style={{ padding: '15px' }}>User Entity</th>
                          <th>Risk Level</th>
                          <th>Risk Score</th>
                          <th>Activity Count</th>
                          <th style={{ color: 'var(--primary)' }}>Files Accessible</th>
                          <th>Security Flags</th>
                        </tr>
                      </thead>
                      <tbody>
                        {riskStats.map((profile, i) => (
                          <tr key={i} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}>
                            <td style={{ padding: '15px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                <div style={{ width: '30px', height: '30px', background: 'rgba(255,255,255,0.05)', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '10px' }}>
                                  {profile.user[0].toUpperCase()}
                                </div>
                                <div>
                                  <div style={{ fontWeight: '600' }}>{profile.user.split('@')[0]}</div>
                                  <div style={{ fontSize: '0.7rem', opacity: 0.6 }}>{profile.user}</div>
                                </div>
                              </div>
                            </td>
                            <td>
                              <span style={{ 
                                padding: '4px 8px', 
                                borderRadius: '4px', 
                                fontSize: '0.7rem', 
                                fontWeight: '800',
                                background: profile.level === 'Critical' ? 'rgba(245, 34, 45, 0.1)' : profile.level === 'Moderate' ? 'rgba(250, 173, 20, 0.1)' : 'rgba(82, 196, 26, 0.1)',
                                color: profile.level === 'Critical' ? '#f5222d' : profile.level === 'Moderate' ? '#faad14' : '#52c41a'
                              }}>
                                {profile.level}
                              </span>
                            </td>
                            <td>
                               <div style={{ width: '100%', background: 'rgba(255,255,255,0.05)', height: '6px', borderRadius: '3px', marginTop: '5px' }}>
                                  <div style={{ width: `${profile.score}%`, height: '100%', background: profile.score > 70 ? '#f5222d' : '#00f2ff', borderRadius: '3px' }}></div>
                               </div>
                               <div style={{ fontSize: '0.7rem', marginTop: '2px' }}>{profile.score}/100</div>
                            </td>
                            <td>{profile.activityCount}</td>
                            <td style={{ fontWeight: '700', color: 'var(--primary)' }}>{profile.fileCount}</td>
                            <td>
                              <div style={{ display: 'flex', gap: '5px', flexWrap: 'wrap' }}>
                                {profile.flags.map((flag, j) => (
                                  <span key={j} style={{ fontSize: '0.65rem', background: 'rgba(255,255,255,0.05)', padding: '2px 6px', borderRadius: '2px' }}>{flag}</span>
                                ))}
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {selectedUser && activeTab !== 'users' && (
                <div className="modal-overlay" onClick={() => setSelectedUser(null)}>
                  <div className="modal-content glass-panel profile-modal" onClick={e => e.stopPropagation()}>
                    <div className="profile-header">
                      <div className="profile-avatar"><UserIcon size={40} /></div>
                      <div className="profile-title-set">
                        <h3>{fullProfile?.profile?.displayName || selectedUser.split('@')[0]}</h3>
                        <p>{selectedUser}</p>
                      </div>
                    </div>
                    
                    <div className="profile-details-grid">
                      <div className="detail-item">
                        <label>Department</label>
                        <span>{fullProfile?.profile?.department || 'Not Specified'}</span>
                      </div>
                      <div className="detail-item">
                        <label>Job Title</label>
                        <span>{fullProfile?.profile?.jobTitle || 'Unassigned'}</span>
                      </div>
                    </div>

                    {fullProfile ? (
                      fullProfile.error ? (
                          <div className="error-box">
                            <ShieldAlert size={20} />
                            <p>Graph API Error: Access Denied. Verify "User.Read.All" & "Files.Read.All" permissions.</p>
                          </div>
                      ) : (
                        <>
                          <div className="stats-row">
                            <div className="stat-pill">
                              <Database size={16} />
                              <span>{fullProfile.stats.uniqueFileCount} Unique Files Touched</span>
                            </div>
                            <div className="stat-pill">
                              <Activity size={16} />
                              <span>{fullProfile.stats.totalEvents} Total Activities</span>
                            </div>
                          </div>

                          <div className="storage-box">
                            <div className="storage-info">
                              <h3>{(fullProfile.storage.used / 1073741824).toFixed(2)} GB</h3>
                              <p>OneDrive Usage</p>
                            </div>
                            <div className="storage-bar-bg">
                              <div className="storage-bar-fill" style={{ width: `${(fullProfile.storage.used / fullProfile.storage.total * 100)}%` }}></div>
                            </div>
                            <p className="storage-footer">Capacity: {(fullProfile.storage.total / 1073741824).toFixed(0)} GB</p>
                          </div>
                        </>
                      )
                    ) : <p className="loading-text">Mining Employee Intelligence Data...</p>}
                    
                    <button className="close-btn" onClick={() => setSelectedUser(null)}>Close Profile</button>
                  </div>
                </div>
              )}
            </>
          )}
        </main>
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <div className="login-page-wrapper">
          {/* Animated background orbs */}
          <div className="login-orb login-orb-1" />
          <div className="login-orb login-orb-2" />
          <div className="login-orb login-orb-3" />

          {/* Grid overlay */}
          <div className="login-grid-overlay" />

          <div className="login-card-outer">
            {/* Left branding panel */}
            <div className="login-brand-panel">
              <div className="login-brand-logo">
                <img src="/ldp-logo.png" alt="LDP Logistics" style={{ width: '90px', height: '90px', objectFit: 'contain', filter: 'brightness(0) invert(1)' }} />
              </div>
              <div className="login-brand-tag">LDP LOGISTICS</div>
              <h1 className="login-brand-title">Data Governance<br/>Command Portal</h1>
              <p className="login-brand-sub">Enterprise-grade Microsoft 365 audit intelligence, real-time file tracking, and insider threat detection — all in one terminal.</p>

              <div className="login-brand-stats">
                <div className="login-stat">
                  <span className="login-stat-val">4,800+</span>
                  <span className="login-stat-label">Events Tracked</span>
                </div>
                <div className="login-stat-divider"/>
                <div className="login-stat">
                  <span className="login-stat-val">42</span>
                  <span className="login-stat-label">Active Users</span>
                </div>
                <div className="login-stat-divider"/>
                <div className="login-stat">
                  <span className="login-stat-val">Live</span>
                  <span className="login-stat-label">Real-time Sync</span>
                </div>
              </div>
            </div>

            {/* Right login panel */}
            <div className="login-form-panel">
              <div className="login-form-inner">
                <div className="login-ms-badge">
                  <svg width="20" height="20" viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>
                  <span>Microsoft Entra ID</span>
                </div>

                <h2 className="login-form-title">Welcome back</h2>
                <p className="login-form-sub">Sign in with your organizational account to access the security terminal.</p>

                <button
                  id="login-btn-microsoft"
                  onClick={handleLogin}
                  disabled={inProgress !== "none"}
                  className="login-ms-btn"
                >
                  {inProgress !== "none" ? (
                    <>
                      <span className="login-spinner"/>
                      Authenticating...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 21 21"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>
                      Login with Microsoft
                    </>
                  )}
                </button>

                {authError && (
                  <div className="login-error-box">
                    <ShieldAlert size={16}/>
                    <span>{authError}</span>
                  </div>
                )}

                <div className="login-divider"><span>secured by</span></div>

                <div className="login-security-pills">
                  <span className="login-sec-pill"><Lock size={11}/> Zero Trust</span>
                  <span className="login-sec-pill"><ShieldAlert size={11}/> MFA Enforced</span>
                  <span className="login-sec-pill"><Activity size={11}/> Audit Logged</span>
                </div>

                <p className="login-footer-note">Restricted to authorized LDP Logistics IT personnel only. All access is monitored and recorded.</p>
              </div>
            </div>
          </div>
        </div>
      </UnauthenticatedTemplate>
    </div>
  );
}

export default App;
