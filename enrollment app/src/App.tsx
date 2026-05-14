import { useState } from 'react';
import { BrowserRouter, Navigate, NavLink, Route, Routes } from 'react-router-dom';
import { ClipboardCheck, Home, Menu } from 'lucide-react';

import { DashboardHomePage } from './pages/DashboardHomePage';
import { SupervisorApprovalPage } from './pages/SupervisorApprovalPage';
import { EnrolmentDetailsPage } from './pages/EnrolmentDetailsPage';
import { EnrolmentCalculationPage } from './pages/EnrolmentCalculationPage';
import { RoleProvider, useRole, ALL_ROLES, ROLE_LABELS, type AppRole } from './context/RoleContext';

const SUPERVISOR_APPROVAL_ROLES: AppRole[] = ['SystemAdmin', 'Supervisor'];
const CALCULATION_ROLES: AppRole[] = ['SystemAdmin', 'Supervisor', 'ENAdmin', 'Verifier'];

function ProtectedRoute({ children, allowedRoles }: { children: React.ReactNode; allowedRoles: AppRole[] }) {
  const { activeRole } = useRole();
  if (!allowedRoles.includes(activeRole)) {
    return <Navigate to="/dashboard-home" replace />;
  }
  return <>{children}</>;
}

function RoleSwitcher({ collapsed }: { collapsed: boolean }) {
  const { activeRole, setActiveRole } = useRole();
  // TODO: gate visibility to SystemAdmin only once real security is implemented
  return (
    <div className={`role-switcher${collapsed ? ' role-switcher--collapsed' : ''}`}>
      {collapsed ? (
        <span className="role-switcher-badge" title={`Acting as: ${ROLE_LABELS[activeRole]}`}>
          {activeRole.slice(0, 2).toUpperCase()}
        </span>
      ) : (
        <>
          <label className="role-switcher-label" htmlFor="role-select">Acting as</label>
          <select
            id="role-select"
            className="role-switcher-select"
            value={activeRole}
            onChange={e => setActiveRole(e.target.value as typeof activeRole)}
          >
            {ALL_ROLES.map(role => (
              <option key={role} value={role}>{ROLE_LABELS[role]}</option>
            ))}
          </select>
        </>
      )}
    </div>
  );
}

function SideNav({ collapsed, onToggle }: { collapsed: boolean; onToggle: () => void }) {
  const { activeRole } = useRole();
  return (
    <aside className={`side-nav${collapsed ? ' collapsed' : ''}`}>
      <button className="side-nav-toggle" type="button" onClick={onToggle} aria-label={collapsed ? 'Expand navigation' : 'Collapse navigation'}>
        <Menu size={24} />
      </button>

      <nav className="side-nav-links" aria-label="Primary">
        <NavLink to="/dashboard-home" className={({ isActive }) => `side-nav-link${isActive ? ' active' : ''}`}>
          <Home size={22} />
          {!collapsed && <span>Dashboard</span>}
        </NavLink>

        {SUPERVISOR_APPROVAL_ROLES.includes(activeRole) && (
          <NavLink to="/supervisor-approval" className={({ isActive }) => `side-nav-link${isActive ? ' active' : ''}`}>
            <ClipboardCheck size={22} />
            {!collapsed && <span>Supervisor Approval</span>}
          </NavLink>
        )}
      </nav>

      <RoleSwitcher collapsed={collapsed} />
    </aside>
  );
}

function AppShell() {
  const [navCollapsed, setNavCollapsed] = useState(false);

  return (
    <div className="app-shell">
      <SideNav collapsed={navCollapsed} onToggle={() => setNavCollapsed(prev => !prev)} />
      <main className="app-shell-content">
        <Routes>
          <Route path="/" element={<Navigate to="/dashboard-home" replace />} />
          <Route path="/dashboard-home" element={<DashboardHomePage />} />
          <Route path="/enrolment/:source/:enrolmentId" element={<EnrolmentDetailsPage />} />
          <Route path="/supervisor-approval" element={<ProtectedRoute allowedRoles={SUPERVISOR_APPROVAL_ROLES}><SupervisorApprovalPage /></ProtectedRoute>} />
          <Route path="/calculation/:source/:enrolmentId" element={<ProtectedRoute allowedRoles={CALCULATION_ROLES}><EnrolmentCalculationPage /></ProtectedRoute>} />
          <Route path="/calculation" element={<Navigate to="/dashboard-home" replace />} />
          <Route path="*" element={<Navigate to="/dashboard-home" replace />} />
        </Routes>
      </main>
    </div>
  );
}

function App() {
  return (
    <BrowserRouter>
      <RoleProvider>
        <AppShell />
      </RoleProvider>
    </BrowserRouter>
  );
}

export default App;
