import React, { useState, useEffect, useMemo } from 'react';
import { Activity, Users, FileText, Share2, LogIn, LayoutDashboard, Database, User as UserIcon, ShieldAlert, Laptop, Globe } from 'lucide-react';
import { ResponsiveSankey } from '@nivo/sankey';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import './App.css';

const SystemClock = () => {
    const [time, setTime] = useState(new Date());
    useEffect(() => {
        const timer = setInterval(() => setTime(new Date()), 1000);
        return () => clearInterval(timer);
    }, []);

    const dateStr = time.toLocaleDateString('en-US', { timeZone: 'America/New_York', year: 'numeric', month: 'numeric', day: 'numeric' });
    const timeStr = time.toLocaleTimeString('en-US', { timeZone: 'America/New_York', hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true }).toLowerCase();

    return (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', color: '#00f3ff', borderRight: '2px solid rgba(0, 243, 255, 0.3)', paddingRight: '15px', marginRight: '15px' }}>
            <span style={{ fontSize: '1.2rem', fontWeight: '800', letterSpacing: '1px' }}>{timeStr}</span>
            <span style={{ fontSize: '0.7rem', color: '#8892b0', textTransform: 'uppercase' }}>{dateStr} (NY/NJ TIME)</span>
        </div>
    );
};
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
              <BarChart data={activityData} margin={{ top: 20, right: 20, bottom: 20, left: -20 }}>
              <defs>
                <linearGradient id="colorEvents" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#ff00f3" stopOpacity={0.8}/>
                  <stop offset="95%" stopColor="#ff00f3" stopOpacity={0}/>
                </linearGradient>
              </defs>
              <XAxis dataKey="hour" stroke="#8892b0" fontSize={12} tickLine={false} axisLine={false} />
              <YAxis stroke="#8892b0" fontSize={12} tickLine={false} axisLine={false} />
              <Tooltip 
                contentStyle={{background: 'rgba(20, 25, 40, 0.9)', border: '1px solid #ff00f3', borderRadius: '8px'}}
                itemStyle={{color: '#ff00f3'}}
              />
              <Bar dataKey="events" fill="url(#colorEvents)" radius={[6, 6, 0, 0]} />
              </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
});

const DashboardContent = React.memo(({ data, error, handleUserClick, sankeyData, activityData, searchTerm }) => {
  // Compute Metrics
  const metrics = useMemo(() => {
    if (!data || data.length === 0) return { total: 0, users: 0, shared: 0 };
    const users = new Set(data.map(d => d.UserId));
    const shared = data.filter(d => d.Operation?.includes('Share')).length;
    return { total: data.length, users: users.size, shared };
  }, [data]);

  const formatNJTime = (creationTime) => {
    if (!creationTime) return "N/A";
    const date = new Date(creationTime);
    // Format: 4/14/2026 time 1:21 pm
    const dateStr = date.toLocaleDateString('en-US', { timeZone: 'America/New_York', year: 'numeric', month: 'numeric', day: 'numeric' });
    const timeStr = date.toLocaleTimeString('en-US', { timeZone: 'America/New_York', hour: 'numeric', minute: '2-digit', hour12: true }).toLowerCase();
    return `${dateStr} time ${timeStr}`;
  };

  const liveFeedData = useMemo(() => {
    if (!data || data.length === 0) return [];
    let processed = [...data].sort((a,b) => new Date(b.CreationTime).getTime() - new Date(a.CreationTime).getTime());
    
    // Search Filter
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
    <>
      <div className="metrics-grid">
        <div className="glass-panel metric-card" style={{ borderLeftColor: "#00f3ff" }}>
          <div className="metric-icon"><Activity size={32} /></div>
          <div className="metric-info">
            <h3>Total File Activity</h3>
            <p>{metrics.total}</p>
          </div>
        </div>
        <div className="glass-panel metric-card" style={{ borderLeftColor: "#ff00f3" }}>
          <div className="metric-icon"><Users size={32} color="#ff00f3" style={{ background: "rgba(255, 0, 243, 0.1)"}} /></div>
          <div className="metric-info">
            <h3>Active Target Users</h3>
            <p>{metrics.users}</p>
          </div>
        </div>
        <div className="glass-panel metric-card" style={{ borderLeftColor: "#00ff88" }}>
          <div className="metric-icon"><Share2 size={32} color="#00ff88" style={{ background: "rgba(0, 255, 136, 0.1)"}} /></div>
          <div className="metric-info">
            <h3>Files Shared externally</h3>
            <p>{metrics.shared}</p>
          </div>
        </div>
      </div>

      <SecurityCharts sankeyData={sankeyData} activityData={activityData} error={error} />

      <div className="glass-panel feed-panel">
        <h2>Live File Activity Matrix (Real-Time Feed)</h2>
        <div className="table-container">
          <table className="feed-table">
            <thead>
              <tr>
                <th>Time</th>
                <th>Active User</th>
                <th>Security Action</th>
                <th>Target Element (File / Path)</th>
                <th>Platform Storage</th>
              </tr>
            </thead>
            <tbody>
              {liveFeedData.map((row) => (
                <tr key={row.id}>
                  <td style={{ color: '#8892b0' }}>{row.time}</td>
                  <td className="user-link" onClick={() => handleUserClick(row.rawEmail)} style={{ fontWeight: 500 }}>{row.user}</td>
                  <td>
                    <span style={{ 
                      padding: '4px 10px', 
                      borderRadius: '12px', 
                      fontSize: '0.8rem',
                      background: row.action.includes('Delete') ? 'rgba(255, 0, 0, 0.2)' : (row.action.includes('Share') ? 'rgba(0, 255, 136, 0.15)' : 'rgba(0, 243, 255, 0.15)'),
                      color: row.action.includes('Delete') ? '#ff4d4d' : (row.action.includes('Share') ? '#00ff88' : '#00f3ff')
                    }}>
                      {row.action}
                    </span>
                  </td>
                  <td>{row.target}</td>
                  <td>{row.location}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
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
  const [activeTab, setActiveTab] = useState('feed'); // feed, web
  const [searchTerm, setSearchTerm] = useState("");

  const AUTHORIZED_EMAILS = [
    'kundan@ldplogistics.com',
    'help-desk@ldplogistics.com', // Original user
    accounts[0]?.username?.toLowerCase()
  ];

  const isAuthorized = AUTHORIZED_EMAILS.includes(accounts[0]?.username?.toLowerCase());

  // Handle Redirect Result
  useEffect(() => {
    instance.handleRedirectPromise().then((response) => {
      if (response) {
        setAuthError(null);
      }
    }).catch(err => {
      console.error("Auth Loop Error:", err);
      setAuthError(err.errorMessage || err.message || "Unknown Security Error");
    });
  }, [instance]);

  const API_BASE = import.meta.env.VITE_API_URL || 'http://localhost:3001';

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


  // Compute Sankey Data
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

  // Compute Activity Data
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
    instance.loginRedirect(loginRequest).catch(e => console.error(e));
  };

  const handleUserClick = async (email) => {
    if (!email || email === 'System') return;
    let fullEmail = email.includes('@') ? email : `${email}@ldplogistics.com`;
    setSelectedUser(fullEmail);
    setFullProfile(null);
    try {
        // Fetch Profile & Stats
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

  return (
    <div className="dashboard-container">
      <AuthenticatedTemplate>
        <header className="header" style={{ marginBottom: '1.5rem' }}>
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <SystemClock />
            <div>
              <h1>Data Governance Dashboard Pro</h1>
              <div style={{ display: 'flex', alignItems: 'center', gap: '15px', marginTop: '4px' }}>
                <p style={{color: '#00ff88', margin: 0, fontSize: '0.8rem', fontWeight: 'bold'}}>ADMIN ACCESS: {accounts[0]?.name}</p>
                <button 
                  onClick={() => instance.logoutRedirect()}
                  style={{ 
                    background: 'rgba(255, 77, 77, 0.15)', 
                    border: '1px solid #ff4d4d', 
                    color: '#ff4d4d', 
                    fontSize: '0.7rem', 
                    padding: '2px 10px', 
                    borderRadius: '12px',
                    cursor: 'pointer'
                  }}
                >
                  Logout
                </button>
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
              <div className="search-wrapper" style={{ position: 'relative' }}>
                <input 
                  type="text" 
                  placeholder="Search file, user or action..." 
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  style={{
                    background: 'rgba(255,255,255,0.05)',
                    border: '1px solid rgba(255,255,255,0.1)',
                    padding: '10px 15px',
                    borderRadius: '20px',
                    color: '#fff',
                    width: '300px',
                    fontSize: '0.9rem',
                    outline: 'none',
                    transition: 'all 0.3s'
                  }}
                  onFocus={(e) => e.target.style.borderColor = '#00f3ff'}
                  onBlur={(e) => e.target.style.borderColor = 'rgba(255,255,255,0.1)'}
                />
              </div>
              <div style={{ color: "#00ff88", display: "flex", alignItems: "center", gap: "10px", fontSize: '0.85rem', fontWeight: "700" }}>
                  <span className="live-dot"></span>
                  LIVE FEED ACTIVE
              </div>
          </div>
        </header>

        {!isAuthorized ? (
          <div className="glass-panel" style={{ padding: '5rem', textAlign: 'center' }}>
             <ShieldAlert size={64} color="#ff4d4d" style={{ marginBottom: '1.5rem' }} />
             <h2 style={{ color: '#fff', fontSize: '2rem' }}>Access Denied</h2>
             <p style={{ color: '#8892b0', margin: '1rem 0' }}>This terminal is restricted to authorized Security Personnel only.</p>
             <p style={{ color: '#ff4d4d', fontWeight: 'bold' }}>Contact the Head of IT to request access for {accounts[0]?.username}.</p>
          </div>
        ) : (
          <>
            <div className="glass-panel" style={{ display: 'flex', gap: '5px', padding: '5px', marginBottom: '2rem', width: 'fit-content' }}>
              <button 
                onClick={() => setActiveTab('feed')}
                style={{ 
                  display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 20px', borderRadius: '12px', border: 'none', cursor: 'pointer',
                  background: activeTab === 'feed' ? 'rgba(0, 243, 255, 0.15)' : 'transparent',
                  color: activeTab === 'feed' ? '#00f3ff' : '#8892b0',
                  fontWeight: 600
                }}
              >
                <Activity size={18} /> File Activity Feed
              </button>
              <button 
                onClick={() => setActiveTab('web')}
                style={{ 
                  display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 20px', borderRadius: '12px', border: 'none', cursor: 'pointer',
                  background: activeTab === 'web' ? 'rgba(0, 255, 136, 0.15)' : 'transparent',
                  color: activeTab === 'web' ? '#00ff88' : '#8892b0',
                  fontWeight: 600
                }}
              >
                <Globe size={18} /> Web & Browser Tracking
              </button>
            </div>

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
                        <th>Time (NJ EST)</th>
                        <th>Device Name</th>
                        <th>User / Account</th>
                        <th>Accessed URL / Activity</th>
                      </tr>
                    </thead>
                    <tbody>
                      {webActivity.length > 0 ? webActivity.map((log, idx) => (
                        <tr key={idx}>
                          <td style={{ color: '#8892b0' }}>{formatNJTime(log.Timestamp)}</td>
                          <td style={{ color: '#00f3ff', fontWeight: '600' }}>{log.DeviceName}</td>
                          <td>{log.InitiatingProcessAccountName}</td>
                          <td style={{ maxWidth: '400px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: '#00ff88' }}>
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
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <div className="login-container glass-panel">
          <div className="login-art">
            <LayoutDashboard size={64} color="#00f3ff" />
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
