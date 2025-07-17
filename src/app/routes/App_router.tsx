import {
  Navigate,
  Route,
  HashRouter as Router,
  Routes,
} from "react-router-dom";
import { AppRoutesNamed } from "./app_routes_named";
import RootHandler from "./handlers/Root_handler";
import AuthHandler from "./handlers/auth/Auth_handler";
import OnboardingHandler from "./handlers/onboarding/Onboarding_handler";
import DashboardHandler from "./handlers/dashboard/Dashboard_handler";
import RolesHandler from "./handlers/roles/Roles_handler";

export default function AppRouter() {
  return (
    <Router>
      <Routes>
        <Route
          path={"*"}
          element={<Navigate to={AppRoutesNamed.ROOT} replace />}
        />
        <Route path={AppRoutesNamed.ROOT} element={<RootHandler />} />
        <Route path={AppRoutesNamed.AUTH} element={<AuthHandler />} />
        <Route
          path={AppRoutesNamed.ONBOARDING}
          element={<OnboardingHandler />}
        />
        <Route
          path={`${AppRoutesNamed.ROLES}:dynamicMode`}
          element={<RolesHandler />}
        />
        <Route path={AppRoutesNamed.DASHBOARD} element={<DashboardHandler />} />
      </Routes>
    </Router>
  );
}
