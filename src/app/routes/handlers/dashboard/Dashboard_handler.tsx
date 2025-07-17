import { useEffect, useState } from "react";
import { RouteHandlerState } from "@shared/domain/router_handler_state";
import { AuthService } from "@shared/application/auth_service";
import { Navigate } from "react-router-dom";
import { AppRoutesNamed, AppRoutesNamedFullPath } from "../../app_routes_named";
import DashboardRootView from "@features/dashboard/presentation/root/Dashboard_root_view";

export default function DashboardHandler() {

    const [authState, setAuthState] = useState(RouteHandlerState.LOADING);

    useEffect(() => {
        const checkAuthState = async () => {
            const state = await AuthService.getAuthState();
            setAuthState(state ? RouteHandlerState.MATCH : RouteHandlerState.NOMATCH);
        };

        checkAuthState();
    }, []);

    const verifyRole = () => {
        const roleSelected = AuthService.getRoleState();
        if (roleSelected) {
            return <DashboardRootView />
        }
        return <Navigate to={`${AppRoutesNamedFullPath.ROLES}true`} replace={true} />;
    }

    return (
        (() => {
            switch (authState) {
                case RouteHandlerState.LOADING:
                    return <div>Loading...</div>;
                case RouteHandlerState.NOMATCH:
                    return <Navigate to={AppRoutesNamed.AUTH} replace={true} />;
                case RouteHandlerState.MATCH:
                    return verifyRole();

            }
        })()
    );
}