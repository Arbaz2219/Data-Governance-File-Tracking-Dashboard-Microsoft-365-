import React, { useState, useEffect, useMemo } from 'react';
import { Activity, Users, FileText, Share2, LogIn, LayoutDashboard, Database, User as UserIcon, ShieldAlert, Laptop, Globe } from 'lucide-react';
import { ResponsiveSankey } from '@nivo/sankey';
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
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
        <div className="glass-panel" style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', color: 'var(--primary)', padding: '6px 15px', marginRight: '15px' }}>
            <span style={{ fontSize: '1.2rem', fontWeight: '800', letterSpacing: '1px', fontFamily: 'Outfit' }}>{timeStr}</span>
            <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '1px' }}>{dateStr} (LOCAL SYSTEM TIME)</span>
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
        <div className="chart-container" style={{ height: "420px", width: "100%" }}>
          {error ? <div style={{color: 'red', marginTop: '20px'}}>{error}</div> : (
            <ResponsiveSankey
                data={sankeyData}
                margin={{ top: 20, right: 180, bottom: 20, left: 180 }}
                align="justify"
                colors={{ scheme: 'category10' }}
                nodeOpacity={1}
                nodeThickness={18}
                nodeSpacing={24}
                nodeBorderWidth={0}
                linkOpacity={0.4}
                linkHoverOpacity={0.7}
                labelPadding={20}
                labelTextColor="#ffffff"
                theme={{
                  labels: { text: { fontSize: 12, fontWeight: 600, fontFamily: 'Outfit, Inter'} },
                  tooltip: { container: { background: "rgba(10, 11, 16, 0.95)", color: "#fff", borderRadius: '8px', border: "1px solid rgba(255,255,255,0.1)" } }
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

const DashboardContent = React.memo(({ data, error, handleUserClick, sankeyData, activityData, searchTerm }) => {
  const metrics = useMemo(() => {
    if (!data || data.length === 0) return { total: 0, users: 0, shared: 0 };
    const users = new Set(data.map(d => d.UserId));
    const shared = data.filter(d => d.Operation?.includes('Share')).length;
    return { total: data.length, users: users.size, shared };
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
      <div className="metrics-grid">
        <div className="metric-card blue-purple">
          <div className="icon-placeholder"><Activity size={20} /></div>
          <div>
            <h3 style={{ margin: '0 0 5px 0', fontSize: '1rem', fontWeight: '800' }}>{metrics.total}</h3>
            <p style={{ margin: 0, fontSize: '0.75rem', opacity: 0.9 }}>Total File Activity</p>
          </div>
          <div style={{ position: 'absolute', right: '15px', bottom: '15px', opacity: 0.2 }}><Activity size={60} /></div>
        </div>
        
        <div className="metric-card cyan">
          <div className="icon-placeholder"><Users size={20} /></div>
          <div>
            <h3 style={{ margin: '0 0 5px 0', fontSize: '1rem', fontWeight: '800' }}>{metrics.users}</h3>
            <p style={{ margin: 0, fontSize: '0.75rem', opacity: 0.9 }}>Active Target Users</p>
          </div>
          <div style={{ position: 'absolute', right: '15px', bottom: '15px', opacity: 0.2 }}><Users size={60} /></div>
        </div>

        <div className="metric-card violet">
          <div className="icon-placeholder"><Share2 size={20} /></div>
          <div>
            <h3 style={{ margin: '0 0 5px 0', fontSize: '1rem', fontWeight: '800' }}>{metrics.shared}</h3>
            <p style={{ margin: 0, fontSize: '0.75rem', opacity: 0.9 }}>External Sharing</p>
          </div>
          <div style={{ position: 'absolute', right: '15px', bottom: '15px', opacity: 0.2 }}><Share2 size={60} /></div>
        </div>
      </div>

      <SecurityCharts sankeyData={sankeyData} activityData={activityData} error={error} />

      <div className="glass-panel feed-panel" style={{ marginTop: '2rem' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 10px' }}>
          <h2 style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
            <span className="status-pulse success"></span>
            Live File Activity Matrix
          </h2>
          <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontWeight: '600' }}>REAL-TIME FEED ACTIVE</span>
        </div>
        <div className="table-container" style={{ padding: '10px' }}>
          <table className="custom-table" style={{ width: '100%' }}>
            <thead>
              <tr style={{ textAlign: 'left', color: 'var(--text-muted)', fontSize: '0.85rem', textTransform: 'uppercase', letterSpacing: '1px' }}>
                <th style={{ padding: '15px' }}>Time</th>
                <th style={{ padding: '15px' }}>Active User</th>
                <th style={{ padding: '15px' }}>Action</th>
                <th style={{ padding: '15px' }}>Target Element</th>
                <th style={{ padding: '15px' }}>Cloud Location</th>
              </tr>
            </thead>
            <tbody>
              {liveFeedData.map((row) => (
                <tr key={row.id}>
                  <td style={{ color: 'var(--text-muted)', fontSize: '0.9rem' }}>{row.time}</td>
                  <td className="user-link" onClick={() => handleUserClick(row.rawEmail)} style={{ color: 'var(--primary)', fontWeight: '600', cursor: 'pointer' }}>{row.user}</td>
                  <td>
                    <span style={{ 
                      padding: '6px 14px', 
                      borderRadius: '20px', 
                      fontSize: '0.75rem',
                      fontWeight: '700',
                      textTransform: 'uppercase',
                      background: row.action.includes('Delete') ? 'rgba(255, 51, 102, 0.15)' : (row.action.includes('Share') ? 'rgba(0, 255, 136, 0.15)' : 'rgba(0, 243, 255, 0.15)'),
                      color: row.action.includes('Delete') ? 'var(--error)' : (row.action.includes('Share') ? 'var(--success)' : 'var(--primary)')
                    }}>
                      {row.action}
                    </span>
                  </td>
                  <td style={{ maxWidth: '300px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{row.target}</td>
                  <td style={{ color: 'var(--secondary)', fontWeight: '500' }}>{row.location}</td>
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

  const AUTHORIZED_EMAILS = [
    'kundan@ldplogistics.com',
    'help-desk@ldplogistics.com'
  ];

  const isAuthorized = AUTHORIZED_EMAILS.includes(accounts[0]?.username?.toLowerCase());

  useEffect(() => {
    if (accounts.length > 0) {
      setAuthError(null);
    }
  }, [accounts]);

  const API_BASE = import.meta.env.VITE_API_URL || 
                   (window.location.hostname === 'localhost' || window.location.hostname.startsWith('192.168') || window.location.hostname.startsWith('172.') || window.location.hostname.startsWith('10.') ? `http://${window.location.hostname}:3001` : window.location.origin);

  const fetchData = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/audit-logs`);
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
      const response = await fetch(`${API_BASE}/api/security/web-activity`);
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
  }, [accounts, activeTab]);

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
        const profileResponse = await fetch(`${API_BASE}/api/user/${fullEmail}/profile`);
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
        const res = await fetch(`${API_BASE}/api/device/${deviceId}/reboot`, { method: 'POST' });
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
            <LayoutDashboard size={32} color="var(--primary)" />
            <h2>SMARTNET</h2>
          </div>
          
          <nav className="nav-menu">
            <div className="nav-label" style={{ color: 'var(--text-muted)', fontSize: '0.7rem', textTransform: 'uppercase', marginBottom: '1rem', paddingLeft: '1rem' }}>Menu</div>
            <div 
              className={`nav-item ${activeTab === 'feed' ? 'active' : ''}`} 
              onClick={() => setActiveTab('feed')}
            >
              <LayoutDashboard size={20} /> Dashboard
            </div>
            <div 
              className={`nav-item ${activeTab === 'web' ? 'active' : ''}`} 
              onClick={() => setActiveTab('web')}
            >
              <ShieldAlert size={20} /> Data Security
            </div>
            <div className="nav-item"><Globe size={20} /> Shadow IT</div>
            <div className="nav-item"><Users size={20} /> User Behavior</div>
            
            <div className="nav-label" style={{ color: 'var(--text-muted)', fontSize: '0.7rem', textTransform: 'uppercase', marginBottom: '1rem', marginTop: '2rem', paddingLeft: '1rem' }}>Manage</div>
            <div className="nav-item"><Database size={20} /> Destinations</div>
            <div className="nav-item"><Activity size={20} /> Protection</div>
          </nav>

          <div className="upgrade-card">
            <h3 style={{ margin: '0 0 10px 0', fontSize: '1rem' }}>Go Pro 👑</h3>
            <p style={{ fontSize: '0.75rem', opacity: 0.9, marginBottom: '1rem' }}>Stay Connected with your team</p>
            <button style={{ background: '#fff', color: '#6e45e2', border: 'none', padding: '8px 16px', borderRadius: '10px', fontWeight: '700', cursor: 'pointer', width: '100%' }}>Upgrade Now</button>
          </div>
        </aside>

        <main className="main-content">
          <header className="topbar">
            <div className="search-bar-wrapper">
              <input 
                type="text" 
                placeholder="Type to search..." 
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            
            <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
              <SystemClock />
              <div className="top-icons" style={{ display: 'flex', gap: '15px', color: 'var(--text-muted)' }}>
                <Activity size={20} style={{ cursor: 'pointer' }} />
                <FileText size={20} style={{ cursor: 'pointer' }} />
                <Globe size={20} style={{ cursor: 'pointer' }} />
              </div>
              
              <div className="profile-pill" style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--bg-card)', padding: '5px 12px', borderRadius: '12px', border: '1px solid var(--glass-border)' }}>
                <div style={{ width: '30px', height: '30px', background: 'var(--primary)', borderRadius: '8px', display: 'flex', alignItems: 'center', justifySelf: 'center' }}>
                  <UserIcon size={18} color="#fff" style={{ margin: 'auto' }} />
                </div>
                <span style={{ fontSize: '0.85rem', fontWeight: '600' }}>Admin</span>
              </div>

              <div className="glass-panel" style={{ padding: '8px 12px', border: 'none', background: 'var(--bg-card)', color: "var(--success)", display: "flex", alignItems: "center", gap: "8px", fontSize: '0.75rem', fontWeight: "700" }}>
                  <span className="status-pulse success" style={{ margin: 0 }}></span>
                  CONNECTED
              </div>
              
              <button 
                onClick={() => instance.logoutRedirect()}
                style={{ background: 'rgba(255, 60, 60, 0.1)', border: 'none', color: 'var(--error)', padding: '8px 12px', borderRadius: '8px', fontWeight: '700', cursor: 'pointer', fontSize: '0.75rem' }}
              >
                LOGOUT
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

              {activeTab === 'web' && (
                <div className="glass-panel feed-panel" style={{ minHeight: '600px' }}>
                  <h2>Browser Search & History Tracking (Microsoft Defender)</h2>
                  <div className="table-container">
                    <table className="feed-table">
                      <thead>
                        <tr>
                          <th>Time (Browser Local)</th>
                          <th>Device Name</th>
                          <th>User / Account</th>
                          <th>Accessed URL / Activity</th>
                        </tr>
                      </thead>
                      <tbody>
                        {webActivity && webActivity.error ? (
                          <tr>
                            <td colSpan="4" style={{ textAlign: 'center', padding: '3rem' }}>
                              <div className="error-box" style={{ maxWidth: '600px', margin: '0 auto', background: 'rgba(255, 60, 60, 0.1)', border: '1px solid var(--error)', padding: '2rem', borderRadius: '12px' }}>
                                <ShieldAlert size={48} color="var(--error)" style={{ marginBottom: '1rem' }} />
                                <h3 style={{ color: 'var(--error)', marginBottom: '10px' }}>{webActivity.error}</h3>
                                <p style={{ color: '#fff', fontSize: '0.9rem' }}>{webActivity.details}</p>
                                <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginTop: '15px' }}>
                                  To capture employee web tracking, you must have Microsoft Defender for Endpoint assigned to your tenant and the user's machines must be physically onboarded.
                                </p>
                              </div>
                            </td>
                          </tr>
                        ) : filteredWebActivity.length > 0 ? filteredWebActivity.map((log, idx) => (
                          <tr key={idx}>
                            <td style={{ color: 'var(--text-muted)' }}>{formatNJTime(log.Timestamp)}</td>
                            <td style={{ color: 'var(--primary)', fontWeight: '600' }}>{log.DeviceName}</td>
                            <td>{log.InitiatingProcessAccountName}</td>
                            <td style={{ maxWidth: '400px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: 'var(--success)' }}>
                              {log.RemoteUrl || 'Internal / Local Process'}
                            </td>
                          </tr>
                        )) : (
                          <tr>
                            <td colSpan="4" style={{ textAlign: 'center', padding: '3rem' }}>
                              <div style={{ opacity: 0.5 }}>
                                <Globe size={48} style={{ marginBottom: '1rem' }} />
                                <p>No real-time web activity detected in the last 10 minutes.</p>
                                <p style={{ fontSize: '0.8rem' }}>Browser searches from Chrome & Edge will appear here automatically.</p>
                              </div>
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {selectedUser && (
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
        <div className="login-container glass-panel">
          <div className="login-art">
            <LayoutDashboard size={64} color="var(--primary)" />
          </div>
          <h2>LDP Logistics</h2>
          <h1>Data Governance Portal</h1>
          <p>Enterprise File Tracking & Audit Governance System</p>
          <button className="login-btn" onClick={handleLogin} disabled={inProgress !== "none"}>
            <LogIn size={20} />
            {inProgress !== "none" ? "Authentication in Progress..." : "Login with Microsoft"}
          </button>

          {authError && (
            <div className="error-box" style={{ marginTop: '1.5rem', fontSize: '0.85rem' }}>
              <ShieldAlert size={20} />
              <div>
                <strong>Security Verification Failed:</strong>
                <p style={{ margin: '5px 0' }}>{authError}</p>
                <p style={{ fontSize: '0.75rem', opacity: 0.7 }}>Check Azure "Single-page application" Platform configuration.</p>
              </div>
            </div>
          )}
          <div className="login-footer">
            <ShieldAlert size={14} />
            Restricted to Authorized Admin Personnel Only
          </div>
        </div>
      </UnauthenticatedTemplate>
    </div>
  );
}

export default App;
