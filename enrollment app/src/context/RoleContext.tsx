import { createContext, useContext, useState, type ReactNode } from 'react';

export type AppRole = 'SystemAdmin' | 'Supervisor' | 'ENAdmin' | 'Verifier';

export const ALL_ROLES: AppRole[] = ['SystemAdmin', 'Supervisor', 'ENAdmin', 'Verifier'];

export const ROLE_LABELS: Record<AppRole, string> = {
  SystemAdmin: 'System Admin',
  Supervisor: 'Supervisor',
  ENAdmin: 'EN Admin',
  Verifier: 'Verifier',
};

const STORAGE_KEY = 'dev_simulated_role';

function readStoredRole(): AppRole {
  try {
    const stored = sessionStorage.getItem(STORAGE_KEY);
    if (stored && ALL_ROLES.includes(stored as AppRole)) return stored as AppRole;
  } catch { /* ignore */ }
  return 'SystemAdmin';
}

type RoleContextValue = {
  activeRole: AppRole;
  setActiveRole: (role: AppRole) => void;
};

const RoleContext = createContext<RoleContextValue | null>(null);

export function RoleProvider({ children }: { children: ReactNode }) {
  const [activeRole, setActiveRoleState] = useState<AppRole>(readStoredRole);

  const setActiveRole = (role: AppRole) => {
    try { sessionStorage.setItem(STORAGE_KEY, role); } catch { /* ignore */ }
    setActiveRoleState(role);
    window.location.reload();
  };

  return (
    <RoleContext.Provider value={{ activeRole, setActiveRole }}>
      {children}
    </RoleContext.Provider>
  );
}

export function useRole(): RoleContextValue {
  const ctx = useContext(RoleContext);
  if (!ctx) throw new Error('useRole must be used within a RoleProvider');
  return ctx;
}
