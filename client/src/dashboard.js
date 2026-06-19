/**
 * LDP LOGISTICS SECURITY DASHBOARD - UTILITIES V2.0
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

// 2. Chart.js Global Standard Configs
export const chartColors = {
  severe: '#ff3b3b',
  major: '#ff8c00',
  moderate: '#ffb800',
  minor: '#00d4ff',
  insignificance: '#2a3a55',
  completed: '#00e676',
  inProgress: '#ffb800',
  pending: '#ff3b3b'
};

export const getDonutConfig = (labels, data, colors, centerText = "") => ({
  labels,
  datasets: [{
    data,
    backgroundColor: colors,
    borderColor: '#0d1629',
    borderWidth: 2,
    hoverOffset: 4,
    cutout: '70%'
  }],
  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: '#1a2540',
        titleColor: '#e2e8f0',
        bodyColor: '#e2e8f0',
        borderColor: '#00d4ff',
        borderWidth: 1,
        padding: 10,
        cornerRadius: 4,
        displayColors: false
      }
    },
    animation: {
      duration: 1500,
      easing: 'easeOutQuart'
    }
  }
});

// 3. Activity Feed Simulation/Processing
// Returns a set of dummy activities for the demo/live feel
export const generateActivityFeed = (rawData = []) => {
  if (rawData.length > 0) {
    return rawData.slice(0, 20).map(item => ({
      id: item.Id || Math.random().toString(36).substr(2, 9),
      user: item.UserId?.split('@')[0] || 'System',
      action: item.Operation || 'Access',
      timestamp: new Date(item.CreationTime).toLocaleTimeString(),
      status: item.Operation?.includes('Delete') ? 'alert' : item.Operation?.includes('Share') ? 'warn' : 'ok'
    }));
  }
  
  // Fallback / Initial sample data
  return [
    { id: 1, user: 'j.doe', action: 'Mass File Download', timestamp: '14:02:11', status: 'alert' },
    { id: 2, user: 'a.smith', action: 'External Sharing', timestamp: '14:05:44', status: 'warn' },
    { id: 3, user: 'system', action: 'Backup Completed', timestamp: '14:08:22', status: 'ok' },
    { id: 4, user: 'r.vaughn', action: 'Login Success', timestamp: '14:10:05', status: 'ok' },
    { id: 5, user: 'k.chen', action: 'Sensitive Doc Access', timestamp: '14:12:30', status: 'warn' }
  ];
};

// 4. Initials Getter for Avatar
export const getInitials = (name = "User") => {
  const parts = name.split(' ');
  if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
  return name.substring(0, 2).toUpperCase();
};
