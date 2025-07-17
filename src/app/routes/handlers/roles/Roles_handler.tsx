import { useEffect, useState } from "react";
import { RouteHandlerState } from "@shared/domain/router_handler_state";
import { AuthService } from "@shared/application/auth_service";
import { Navigate, useParams } from "react-router-dom";
import { AppRoutesNamed } from "../../app_routes_named";
import RolesView from "@features/roles/presentation/Roles_view";

export default function RolesHandler() {
  const [authState, setAuthState] = useState(RouteHandlerState.LOADING);

  useEffect(() => {
    const checkAuthState = async () => {
      const state = await AuthService.getAuthState();
      setAuthState(state ? RouteHandlerState.MATCH : RouteHandlerState.NOMATCH);
    };

    checkAuthState();
  }, []);

  const { dynamicMode } = useParams();

  return (() => {
    switch (authState) {
      case RouteHandlerState.LOADING:
        return <div>Loading...</div>;
      case RouteHandlerState.NOMATCH:
        return <Navigate to={AppRoutesNamed.AUTH} replace={true} />;
      case RouteHandlerState.MATCH:
        return <RolesView dynamicMode={dynamicMode === "true"} />;
    }
  })();
}
