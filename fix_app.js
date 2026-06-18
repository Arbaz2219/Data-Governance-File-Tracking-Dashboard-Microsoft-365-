const fs = require('fs');
let c = fs.readFileSync('client/src/App.jsx', 'utf8');

// 1. Fix Super Admin - remove arbaz, add help-desk
c = c.replace(
  "const SUPER_ADMINS = ['kundan@ldplogistics.com', 'arbaz@ldplogistics.com'];",
  "const SUPER_ADMINS = ['kundan@ldplogistics.com', 'help-desk@ldplogistics.com'];"
);

// 2. Update table header - add Dashboard Access column
c = c.replace('<th>Action</th>', '<th>Dashboard Access</th>\r\n                                  <th>Investigate</th>');

// 3. Find the Investigate-only td and replace with Grant/Revoke + Investigate
const marker = 'className="tab-btn" \r\n                                           style={{ padding: \'6px 12px\', background: \'var(--primary)\', color: \'#000\' }}\r\n                                           onClick={() => {\r\n                                              setSelectedUserIntel(u.id);\r\n                                              handleUserClick(u.id);\r\n                                           }}\r\n                                        >\r\n                                           Investigate';

if (c.includes('ShieldX size={12} /> Revoke')) {
  console.log('Grant/Revoke already present, skipping step 3');
} else {
  // Insert Grant/Revoke td BEFORE the existing Investigate td
  const investigateStart = '<td>\r\n                                        <button \r\n                                           className="tab-btn" \r\n                                           style={{ padding: \'6px 12px\', background: \'var(--primary)\', color: \'#000\' }}';
  
  const grantRevokeBlock = `<td>
                                        {adminStatus.isSuperAdmin && (
                                           authorizedUsers.includes(u.id.toLowerCase()) ? (
                                              <button
                                                 style={{ padding: '5px 12px', background: 'rgba(255,77,79,0.12)', color: '#ff4d4f', border: '1px solid rgba(255,77,79,0.3)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', fontSize: '0.7rem', display: 'flex', alignItems: 'center', gap: '5px' }}
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
                                                 style={{ padding: '5px 12px', background: 'rgba(0,245,212,0.1)', color: 'var(--primary)', border: '1px solid rgba(0,245,212,0.3)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', fontSize: '0.7rem', display: 'flex', alignItems: 'center', gap: '5px' }}
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
                                        )}
                                        {!adminStatus.isSuperAdmin && (
                                           <span style={{ fontSize: '0.7rem', color: authorizedUsers.includes(u.id.toLowerCase()) ? '#52c41a' : 'rgba(255,255,255,0.3)' }}>
                                              {authorizedUsers.includes(u.id.toLowerCase()) ? '✓ Authorized' : '✗ Restricted'}
                                           </span>
                                        )}
                                     </td>\r\n                                     ` + investigateStart;

  c = c.replace(investigateStart, grantRevokeBlock);
}

fs.writeFileSync('client/src/App.jsx', c, 'utf8');
const lines = c.split('\n').length;
console.log('Done. Total lines: ' + lines);

// Verify key changes
if (c.includes('help-desk@ldplogistics.com')) console.log('OK: Super Admin fixed');
if (c.includes('Dashboard Access')) console.log('OK: Table header updated');
if (c.includes('Grant Access') || c.includes('ShieldX size')) console.log('OK: Grant/Revoke buttons present');
