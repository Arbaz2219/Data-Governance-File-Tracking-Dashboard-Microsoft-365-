// v1.0.2 - Forcing clean build to resolve potential cache issues
import React, { useState, useEffect, useMemo } from 'react';
import { Activity, Users, FileText, Share2, LogIn, LayoutDashboard, Database, User as UserIcon, ShieldAlert, Laptop, Globe, Lock, Shield, Settings, UserPlus, UserMinus, ShieldCheck, ShieldX, Search, Plus, Trash2, AlertTriangle } from 'lucide-react';
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
        const term = String(searchTerm).toLowerCase();
        processed = processed.filter(d => 
            String(d.UserId || '').toLowerCase().includes(term) || 
            String(d.ObjectId || '').toLowerCase().includes(term) || 
            String(d.Operation || '').toLowerCase().includes(term)
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
    <div className="admin-portal-container">
      <div className="glass-panel" style={{ padding: '30px', maxWidth: '800px', margin: '0 auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px', marginBottom: '30px' }}>
          <ShieldCheck size={32} color="var(--secondary)" />
          <h2 style={{ margin: 0, fontSize: '1.5rem' }}>Access Control Terminal</h2>
        </div>

        <div className="add-user-form" style={{ display: 'flex', gap: '10px', marginBottom: '40px' }}>
          <input 
            type="email" 
            placeholder="Enter Corporate Email to Authorize..." 
            value={newEmail}
            onChange={(e) => setNewEmail(e.target.value)}
            style={{ flex: 1, padding: '12px', background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', borderRadius: '8px', color: '#fff' }}
          />
          <button 
            onClick={handleAdd} 
            disabled={loading}
            style={{ padding: '0 25px', background: 'var(--secondary)', color: '#000', border: 'none', borderRadius: '8px', fontWeight: '700', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px' }}
          >
            <UserPlus size={18} /> {loading ? 'Processing...' : 'Grant Access'}
          </button>
        </div>

        <div className="users-auth-list">
          <h3 style={{ fontSize: '0.9rem', color: 'var(--text-muted)', marginBottom: '15px' }}>Currently Authorized Identities</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
            {authorizedUsers.map(email => (
              <div key={email} className="auth-user-row" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '15px', background: 'rgba(255,255,255,0.02)', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.05)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <div style={{ width: '8px', height: '8px', borderRadius: '50%', background: email === 'kundan@ldplogistics.com' ? 'var(--secondary)' : 'var(--success)' }}></div>
                  <span>{email}</span>
                  {email === 'kundan@ldplogistics.com' && <span style={{ fontSize: '0.6rem', background: 'var(--secondary)', color: '#000', padding: '2px 6px', borderRadius: '4px', fontWeight: '800' }}>SUPER ADMIN</span>}
                </div>
                {email !== 'kundan@ldplogistics.com' && (
                  <button onClick={() => handleRemove(email)} className="revoke-btn" style={{ background: 'transparent', border: 'none', color: '#ff4d4f', cursor: 'pointer', padding: '5px' }}>
                    <Trash2 size={18} />
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

function App() {
  const { instance, accounts, inProgress } = useMsal();
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [authorizedUsers, setAuthorizedUsers] = useState([]);
  const [authLoading, setAuthLoading] = useState(true);
  const [newAuthEmail, setNewAuthEmail] = useState('');
  const [adminStatus, setAdminStatus] = useState({ isSuperAdmin: false, isAuthorized: false });
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

  const SUPER_ADMINS = ['help-desk@ldplogistics.com', 'kundan@ldplogistics.com'];

  const toggleTheme = () => {
    setIsDarkMode(!isDarkMode);
    document.body.classList.toggle('dark-mode');
  };

  const fetchAuthList = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/admin/authorized-users`, {
        headers: { 'X-Dashboard-Key': 'LDP_SECURE_9821_!@#$' }
      });
      const list = await response.json();
      const authorizedList = Array.isArray(list) ? list : [];
      setAuthorizedUsers(authorizedList);
      
      const email = accounts[0]?.username?.toLowerCase();
      const isSuper = SUPER_ADMINS.includes(email);
      setAdminStatus({
        isSuperAdmin: isSuper,
        isAuthorized: authorizedList.includes(email) || isSuper
      });
    } catch (e) {
      console.error("Auth sync failed", e);
      // Fallback for Super Admins even if API is down
      const email = accounts[0]?.username?.toLowerCase();
      if (SUPER_ADMINS.includes(email)) {
        setAdminStatus({ isSuperAdmin: true, isAuthorized: true });
      }
    } finally {
      setAuthLoading(false);
    }
  };

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

  const userIntelligenceSummary = useMemo(() => {
    if (!data || data.length === 0) return [];
    const stats = {};
    data.forEach(log => {
        const userId = log.UserId?.toLowerCase();
        if (!userId || userId.includes('app@sharepoint')) return;
        if (!stats[userId]) {
            stats[userId] = {
                id: userId,
                files: new Set(),
                actions: 0,
                lastSeen: log.CreationTime,
                workload: log.Workload
            };
        }
        stats[userId].actions++;
        if (log.ObjectId) stats[userId].files.add(log.ObjectId);
        if (new Date(log.CreationTime) > new Date(stats[userId].lastSeen)) {
            stats[userId].lastSeen = log.CreationTime;
        }
    });
    return Object.values(stats).sort((a, b) => b.actions - a.actions);
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
    const term = String(searchTerm).toLowerCase();
    return webActivity.filter(log => 
      String(log.DeviceName || '').toLowerCase().includes(term) || 
      String(log.RemoteUrl || '').toLowerCase().includes(term) || 
      String(log.InitiatingProcessAccountName || '').toLowerCase().includes(term)
    );
  }, [webActivity, searchTerm]);

  useEffect(() => {
    if (accounts.length > 0) {
      fetchAuthList();
    }
  }, [accounts]);

  // CLEAN CONSOLIDATED POLLING SYSTEM
  useEffect(() => {
    if (accounts.length > 0 && adminStatus.isAuthorized) {
      // Initial load
      fetchData();
      fetchWebActivity();
      fetchActiveCalls();
      fetchSecurityInsights();
      
      const interval = setInterval(() => {
        fetchData(); // Always refresh main dashboard feed
        
        // Context-aware background refreshing
        if (activeTab === 'web') fetchWebActivity();
        if (activeTab === 'comms') fetchActiveCalls();
        if (activeTab === 'shadow' || activeTab === 'behavior') fetchSecurityInsights();
      }, 5000); 
      
      return () => clearInterval(interval);
    }
  }, [accounts, activeTab, adminStatus.isAuthorized]);

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
        {authLoading ? (
          <div style={{ height: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#0A0F1E' }}>
            <div className="loading-spinner"></div>
            <p style={{ marginLeft: '20px', color: 'var(--primary)', fontWeight: '800' }}>VERIFYING TERMINAL ACCESS...</p>
          </div>
        ) : !adminStatus.isAuthorized ? (
          <div className="access-denied-wrapper" style={{ height: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#0A0F1E', color: '#fff', textAlign: 'center' }}>
            <div className="glass-panel" style={{ padding: '60px', maxWidth: '500px', border: '1px solid #ff4d4f', boxShadow: '0 0 30px rgba(255, 77, 79, 0.2)' }}>
              <Lock size={64} color="#ff4d4f" style={{ marginBottom: '20px' }} />
              <h1 style={{ color: '#ff4d4f', marginBottom: '10px' }}>ACCESS RESTRICTED</h1>
              <p style={{ color: 'var(--text-muted)', lineHeight: '1.6' }}>
                Your identity **{accounts[0]?.username}** is not authorized to access this terminal. 
                Contact the system administrator to request access.
              </p>
              <button 
                onClick={() => instance.logoutRedirect()} 
                style={{ marginTop: '30px', padding: '12px 30px', background: '#ff4d4f', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: '800' }}
              >
                Logout / Switch Account
              </button>
            </div>
          </div>
        ) : (
          <>
            <aside className="sidebar">
          <div className="logo-section">
            <div style={{ padding: '8px', background: 'rgba(20, 184, 166, 0.1)', borderRadius: '10px', display: 'flex' }}>
              <LayoutDashboard size={24} color="var(--primary)" />
            </div>
            <h2>LDP LOGISTICS</h2>
          </div>
          
          <div className="sidebar-nav">
            <button className={`nav-item ${activeTab === 'dashboard' ? 'active' : ''}`} onClick={() => setActiveTab('dashboard')}><LayoutDashboard size={18} /> Overview</button>
            <button className={`nav-item ${activeTab === 'users' ? 'active' : ''}`} onClick={() => setActiveTab('users')}><Users size={18} /> User Intelligence</button>
            <button className={`nav-item ${activeTab === 'behavior' ? 'active' : ''}`} onClick={() => setActiveTab('behavior')}><ShieldAlert size={18} /> Security Pulse</button>
            <button className={`nav-item ${activeTab === 'comms' ? 'active' : ''}`} onClick={() => setActiveTab('comms')}><Activity size={18} /> Comms Intel</button>
            <button className={`nav-item ${activeTab === 'web' ? 'active' : ''}`} onClick={() => setActiveTab('web')}><Globe size={18} /> Web Activity</button>
            <button className={`nav-item ${activeTab === 'shadow' ? 'active' : ''}`} onClick={() => setActiveTab('shadow')}><Shield size={18} /> Logs</button>
            
            {adminStatus.isSuperAdmin && (
              <button 
                className={`nav-item ${activeTab === 'admin' ? 'active' : ''}`} 
                onClick={() => setActiveTab('admin')} 
                style={{ marginTop: '20px', borderTop: '1px solid rgba(255,255,255,0.1)', paddingTop: '20px', color: 'var(--secondary)' }}
              >
                <Settings size={18} /> Admin Portal
              </button>
            )}
          </div>
        </aside>

          <main className="dashboard-main main-content">
            {/* Header Content */}
            <div className="top-header">
              <div className="system-status">
                <div className="status-dot online"></div>
                <span style={{ color: 'var(--success)', fontWeight: '800', fontSize: '0.7rem' }}>SECURE CONNECTION ACTIVE</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                <div className="user-pill">
                  <UserIcon size={14} />
                  <span>{accounts[0]?.name?.split(' ')[0]} {adminStatus.isSuperAdmin ? '(ADMIN)' : ''}</span>
                </div>
                <SystemClock />
                <button onClick={() => instance.logoutRedirect()} className="logout-icon-btn" title="Secure Logout">
                  <LogIn size={20} />
                </button>
              </div>
            </div>

            {/* TAB RENDERING LOGIC */}
            <>
              {activeTab === 'dashboard' && (
                <div className="dashboard-content">
                  {/* TOP KPI CARDS */}
                  <div className="metrics-grid kpi-row" style={{ gridTemplateColumns: 'repeat(4, 1fr)', gap: '1.5rem', marginBottom: '2rem' }}>
                    <div className="glass-panel" style={{ padding: '25px', textAlign: 'center', borderBottom: '3px solid var(--primary)' }}>
                      <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: '800', textTransform: 'uppercase', marginBottom: '10px' }}>Risk Score</div>
                      <div style={{ fontSize: '2.4rem', fontWeight: '800', color: 'var(--primary)' }}>0.0%</div>
                    </div>
                    <div className="glass-panel" style={{ padding: '25px', textAlign: 'center', borderBottom: '3px solid var(--secondary)' }}>
                      <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: '800', textTransform: 'uppercase', marginBottom: '10px' }}>Active Incidents</div>
                      <div style={{ fontSize: '2.4rem', fontWeight: '800', color: 'var(--secondary)' }}>0</div>
                    </div>
                    <div className="glass-panel" style={{ padding: '25px', textAlign: 'center', borderBottom: '3px solid var(--warning)' }}>
                      <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: '800', textTransform: 'uppercase', marginBottom: '10px' }}>Vulnerability Rate</div>
                      <div style={{ fontSize: '2.4rem', fontWeight: '800', color: 'var(--warning)' }}>1.8%</div>
                    </div>
                    <div className="glass-panel" style={{ padding: '25px', textAlign: 'center', borderBottom: '3px solid var(--success)' }}>
                      <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: '800', textTransform: 'uppercase', marginBottom: '10px' }}>Compliance Score</div>
                      <div style={{ fontSize: '2.4rem', fontWeight: '800', color: 'var(--success)' }}>52.6%</div>
                    </div>
                  </div>

                  {/* MIDDLE ROW PANELS */}
                  <div className="analysis-row" style={{ display: 'grid', gridTemplateColumns: '1fr 1.5fr 1fr', gap: '1.5rem', marginBottom: '2rem' }}>
                    <div className="glass-panel" style={{ padding: '20px' }}>
                      <h3 style={{ margin: '0 0 15px 0', fontSize: '1rem' }}>Risk Rating Breakdown</h3>
                      <div style={{ height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'radial-gradient(circle, rgba(0,245,212,0.05) 0%, transparent 70%)' }}>
                        <div style={{ width: '120px', height: '120px', borderRadius: '50%', border: '15px solid var(--primary)', borderRightColor: 'var(--secondary)', borderBottomColor: 'var(--warning)', borderLeftColor: 'var(--success)' }}></div>
                      </div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '10px', marginTop: '15px', fontSize: '0.7rem' }}>
                        <span style={{ color: 'var(--primary)' }}>● Low</span>
                        <span style={{ color: 'var(--secondary)' }}>● High</span>
                        <span style={{ color: 'var(--warning)' }}>● Medium</span>
                        <span style={{ color: 'var(--success)' }}>● Safe</span>
                      </div>
                    </div>

                    <div className="glass-panel" style={{ padding: '20px' }}>
                      <h3 style={{ margin: '0 0 15px 0', fontSize: '1rem' }}>Risk Heat Map</h3>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '5px' }}>
                        {[...Array(25)].map((_, i) => (
                          <div key={i} style={{ 
                            height: '40px', 
                            background: i === 0 || i === 6 || i === 12 ? 'var(--error)' : i < 10 ? 'var(--warning)' : 'var(--success)',
                            opacity: 0.6 + (Math.random() * 0.4),
                            borderRadius: '4px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontSize: '0.7rem',
                            fontWeight: '800'
                          }}>{Math.floor(Math.random() * 9)}</div>
                        ))}
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '10px', fontSize: '0.65rem', color: 'var(--text-muted)' }}>
                        <span>Insignificant</span>
                        <span>Severe</span>
                      </div>
                    </div>

                    <div className="glass-panel" style={{ padding: '20px', textAlign: 'center' }}>
                      <h3 style={{ margin: '0 0 15px 0', fontSize: '1rem' }}>Action Plan Breakdown</h3>
                      <div style={{ position: 'relative', width: '160px', height: '80px', margin: '20px auto 0', overflow: 'hidden' }}>
                        <div style={{ width: '160px', height: '160px', borderRadius: '50%', border: '20px solid rgba(255,255,255,0.05)', borderTopColor: 'var(--primary)', transform: 'rotate(-45deg)' }}></div>
                        <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, fontSize: '1.5rem', fontWeight: '800' }}>72%</div>
                      </div>
                      <p style={{ fontSize: '0.75rem', color: 'var(--text-muted)', marginTop: '10px' }}>Mitigation Tasks Completed</p>
                    </div>
                  </div>

                  {/* LIVE SECURITY FEED TABLE */}
                  <div className="glass-panel" style={{ marginTop: '2rem', padding: '25px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
                        <h3 style={{ margin: 0 }}>🚨 Live Security Feed</h3>
                        <span style={{ fontSize: '0.7rem', color: 'var(--secondary)' }}>REAL-TIME MONITORING ACTIVE</span>
                    </div>
                    <div className="table-container">
                        <table className="custom-table" style={{ width: '100%' }}>
                            <thead>
                                <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.8rem' }}>
                                    <th>Timestamp (EST)</th>
                                    <th>Identity</th>
                                    <th>Operation</th>
                                    <th>Resource Path</th>
                                </tr>
                            </thead>
                            <tbody>
                                {data.length === 0 ? (
                                    <tr>
                                        <td colSpan="4" style={{ textAlign: 'center', padding: '40px', color: 'var(--text-muted)', fontSize: '0.9rem' }}>
                                            <div style={{ marginBottom: '10px' }}><ShieldCheck size={32} opacity={0.3} /></div>
                                            No active threats or transactions detected in the current sync window.
                                        </td>
                                    </tr>
                                ) : data.slice(-15).reverse().map((log, i) => (
                                    <tr key={i}>
                                        <td style={{ fontSize: '0.8rem' }}>{formatNJTime(log.CreationTime)}</td>
                                        <td style={{ fontWeight: '700' }}>{log.UserId?.split('@')[0]}</td>
                                        <td>
                                            <span style={{ 
                                                padding: '4px 8px', 
                                                background: log.Operation === 'FileDeleted' ? 'rgba(255, 77, 79, 0.1)' : 'rgba(0, 245, 212, 0.1)', 
                                                color: log.Operation === 'FileDeleted' ? '#ff4d4f' : 'var(--primary)',
                                                borderRadius: '4px',
                                                fontSize: '0.7rem',
                                                fontWeight: '800'
                                            }}>
                                                {log.Operation}
                                            </span>
                                        </td>
                                        <td style={{ fontSize: '0.75rem', opacity: 0.8, maxWidth: '400px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                            {log.isSensitive && <AlertTriangle size={14} color="#ff4d4f" title="Sensitive Content" style={{ marginRight: '5px', verticalAlign: 'middle' }} />}
                                            {log.ObjectId?.split('/').pop()}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                  </div>
                </div>
              )}

            {activeTab === 'users' && (
              <div className="user-intelligence-container" style={{ animation: 'fadeIn 0.3s ease' }}>
                {!selectedUserIntel ? (
                   <div className="glass-panel" style={{ padding: '30px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
                         <div>
                            <h2 style={{ margin: 0 }}>Employee Intelligence Overview</h2>
                            <p style={{ color: 'var(--text-muted)', fontSize: '0.9rem' }}>Comprehensive monitoring of all active M365 identities</p>
                         </div>
                         <input 
                            type="text" 
                            placeholder="Search employees..." 
                            className="search-input"
                            style={{ width: '300px' }}
                            onChange={(e) => setSearchTerm(e.target.value)}
                         />
                      </div>
                      
                      <div style={{ overflowX: 'auto' }}>
                         <table className="custom-table" style={{ width: '100%' }}>
                            <thead>
                               <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.8rem' }}>
                                  <th>Employee</th>
                                  <th>Risk Level</th>
                                  <th>Sensitivity</th>
                                  <th>Action Count</th>
                                  <th>Files Accessed</th>
                                  <th>Dashboard Access</th>
                                  <th>Investigate</th>
                               </tr>
                            </thead>
                            <tbody>
                               {userIntelligenceSummary.filter(u => String(u.id || '').toLowerCase().includes(String(searchTerm || '').toLowerCase())).map((u, i) => (
                                  <tr key={i}>
                                     <td>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                                           <div style={{ width: '36px', height: '36px', background: 'rgba(0, 245, 212, 0.1)', border: '1px solid var(--primary)', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--primary)', fontWeight: '800' }}>
                                              {u.id[0].toUpperCase()}
                                           </div>
                                           <div>
                                              <div style={{ fontWeight: '700' }}>{u.id.split('@')[0]}</div>
                                              <div style={{ fontSize: '0.7rem', opacity: 0.6 }}>{u.id}</div>
                                           </div>
                                        </div>
                                     </td>
                                     <td>
                                        <span style={{ 
                                            padding: '4px 10px', 
                                            borderRadius: '4px', 
                                            fontSize: '0.7rem', 
                                            fontWeight: '800',
                                            background: riskStats.find(r => String(r.user || '').toLowerCase() === String(u.id || '').toLowerCase())?.level === 'Critical' ? 'rgba(245, 34, 45, 0.1)' : 'rgba(0, 245, 212, 0.1)',
                                            color: riskStats.find(r => String(r.user || '').toLowerCase() === String(u.id || '').toLowerCase())?.level === 'Critical' ? '#ff4d4f' : 'var(--primary)'
                                        }}>
                                            {riskStats.find(r => String(r.user || '').toLowerCase() === String(u.id || '').toLowerCase())?.level || 'LOW'}
                                        </span>
                                     </td>
                                     <td>
                                        {data.some(d => d.UserId?.toLowerCase() === u.id.toLowerCase() && d.isSensitive) ? (
                                           <span style={{ color: '#ff4d4f', display: 'flex', alignItems: 'center', gap: '4px', fontSize: '0.7rem', fontWeight: '800' }}>
                                              <AlertTriangle size={12} /> SENSITIVE
                                           </span>
                                        ) : (
                                           <span style={{ color: 'var(--success)', fontSize: '0.7rem', fontWeight: '800' }}>SECURE</span>
                                        )}
                                     </td>
                                     <td><span style={{ fontWeight: '700', color: 'var(--primary)' }}>{u.actions}</span></td>
                                     <td><span style={{ fontWeight: '700' }}>{u.files.size}</span></td>
                                     <td>
                                        {adminStatus.isSuperAdmin && (
                                           authorizedUsers.includes(u.id.toLowerCase()) ? (
                                              <button
                                                 style={{ padding: '5px 12px', background: 'rgba(255,77,79,0.12)', color: '#ff4d4f', border: '1px solid rgba(255,77,79,0.3)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', fontSize: '0.7rem', display: 'flex', alignItems: 'center', gap: '5px' }}
                                                 onClick={async () => {
                                                    if (!window.confirm(`Revoke access for ${u.id}?`)) return;
                                                    await fetch(`${API_BASE}/api/admin/authorized-users`, {
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
                                                 style={{ padding: '5px 12px', background: 'rgba(0,245,212,0.1)', color: 'var(--primary)', border: '1px solid rgba(0,245,212,0.3)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', fontSize: '0.7rem', display: 'flex', alignItems: 'center', gap: '5px' }}
                                                 onClick={async () => {
                                                    await fetch(`${API_BASE}/api/admin/authorized-users`, {
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
                                        )}
                                        {!adminStatus.isSuperAdmin && (
                                           <span style={{ fontSize: '0.7rem', color: authorizedUsers.includes(u.id.toLowerCase()) ? '#52c41a' : 'var(--text-muted)' }}>
                                              {authorizedUsers.includes(u.id.toLowerCase()) ? 'Authorized' : 'Restricted'}
                                           </span>
                                        )}
                                     </td>
                                     <td>
                                        <button 
                                           className="tab-btn" 
                                           style={{ padding: '6px 12px', background: 'var(--primary)', color: '#000' }}
                                           onClick={() => {
                                              setSelectedUserIntel(u.id);
                                              handleUserClick(u.id);
                                           }}
                                        >
                                           Investigate
                                        </button>
                                     </td>
                                  </tr>
                               ))}
                            </tbody>
                         </table>
                      </div>
                   </div>
                ) : (
                   <div style={{ animation: 'fadeIn 0.5s ease' }}>
                      <div style={{ marginBottom: '1.5rem' }}>
                         <button 
                            className="tab-btn" 
                            style={{ background: 'rgba(255,255,255,0.05)', color: '#fff' }}
                            onClick={() => setSelectedUserIntel(null)}
                         >
                            ← Back to Overview
                         </button>
                      </div>
                      <div style={{ height: 'calc(100vh - 220px)', overflowY: 'auto' }}>
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

                         <div className="glass-panel" style={{ padding: '25px', marginTop: '2rem' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                               <h3 style={{ margin: 0 }}>Recent File Access Details</h3>
                            </div>
                            <div className="table-container" style={{ maxHeight: '400px', overflowY: 'auto' }}>
                               <table className="custom-table" style={{ width: '100%' }}>
                                  <thead style={{ position: 'sticky', top: 0, background: 'var(--bg-dark)', zIndex: 1 }}>
                                     <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.8rem' }}>
                                        <th>Timestamp</th>
                                        <th>Operation</th>
                                        <th>File Name</th>
                                        <th>Sensitivity</th>
                                     </tr>
                                  </thead>
                                  <tbody>
                                     {data.filter(log => log.UserId?.toLowerCase() === selectedUserIntel.toLowerCase() && log.ObjectId).slice(0, 100).map((log, i) => (
                                        <tr key={i}>
                                           <td style={{ fontSize: '0.8rem' }}>{new Date(log.CreationTime).toLocaleString()}</td>
                                           <td style={{ fontSize: '0.8rem', color: 'var(--secondary)' }}>{log.Operation}</td>
                                           <td style={{ fontSize: '0.85rem', wordBreak: 'break-all' }}>{log.ObjectId.split('/').pop()}</td>
                                           <td>
                                              {log.isSensitive ? (
                                                 <span style={{ color: '#ff4d4f', display: 'flex', alignItems: 'center', gap: '4px', fontSize: '0.7rem', fontWeight: '800' }}>
                                                    <AlertTriangle size={12} /> SENSITIVE
                                                 </span>
                                              ) : (
                                                 <span style={{ color: 'var(--success)', fontSize: '0.7rem', fontWeight: '800' }}>SECURE</span>
                                              )}
                                           </td>
                                        </tr>
                                     ))}
                                     {data.filter(log => log.UserId?.toLowerCase() === selectedUserIntel.toLowerCase() && log.ObjectId).length === 0 && (
                                        <tr><td colSpan="4" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-muted)' }}>No recent file access logs found for this user.</td></tr>
                                     )}
                                  </tbody>
                               </table>
                            </div>
                         </div>
                       </div>
                    </div>
                 )}
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


            {activeTab === 'shadow' && (
              <div className="behavior-container" style={{ animation: 'fadeIn 0.5s ease-out', padding: '20px' }}>
                <div className="glass-panel" style={{ padding: '25px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                     <h2 style={{ margin: 0, color: 'var(--primary)', display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <Shield size={24} /> User Login Logs (Entra ID)
                     </h2>
                  </div>
                  <div className="table-container">
                    <table className="custom-table" style={{ width: '100%' }}>
                      <thead>
                        <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.8rem', textTransform: 'uppercase' }}>
                          <th style={{ padding: '15px' }}>Timestamp</th>
                          <th>User</th>
                          <th>Operation</th>
                          <th>IP Address</th>
                          <th>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {data.filter(log => log.Workload === 'AzureActiveDirectory' || log.Operation?.includes('Login') || log.Operation?.includes('Logon') || log.Operation === 'UserLoggedIn').slice(0, 50).map((log, i) => (
                          <tr key={i} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}>
                            <td style={{ padding: '15px' }}>{new Date(log.CreationTime).toLocaleString()}</td>
                            <td>{log.UserId}</td>
                            <td>{log.Operation}</td>
                            <td>{log.ClientIP || 'N/A'}</td>
                            <td>
                               <span style={{ color: log.ResultStatus === 'Failed' ? '#ff4d4f' : 'var(--success)' }}>
                                  {log.ResultStatus || 'Success'}
                               </span>
                            </td>
                          </tr>
                        ))}
                        {data.filter(log => log.Workload === 'AzureActiveDirectory' || log.Operation?.includes('Login') || log.Operation?.includes('Logon') || log.Operation === 'UserLoggedIn').length === 0 && (
                           <tr><td colSpan="5" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-muted)' }}>No Entra ID login logs found.</td></tr>
                        )}
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
            {activeTab === 'admin' && adminStatus.isSuperAdmin && (
              <AdminPortal 
                authorizedUsers={authorizedUsers} 
                fetchAuthList={fetchAuthList} 
                API_BASE={API_BASE} 
              />
            )}
              </main>
            </>
          )}
        </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <div className="login-page-wrapper dark-theme-override">
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
              <div className="login-brand-tag" style={{ color: 'var(--secondary)' }}>DP LOGISTICS</div>
              <h1 className="login-brand-title">Security Risk<br/>Management Portal</h1>
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
