import { useEffect, useState } from "react";
import { AuthService } from "@shared/application/auth_service";
import { RouteHandlerState } from "@shared/domain/router_handler_state";
import { AppRoutesNamed } from "../app_routes_named";
import { Navigate } from "react-router-dom";

export default function RootHandler() {

    const [authState, setAuthState] = useState(RouteHandlerState.LOADING);

    useEffect(() => {
        const checkAuthState = async () => {
            const state = await AuthService.getAuthState();
            setAuthState(state ? RouteHandlerState.MATCH : RouteHandlerState.NOMATCH);
        };

        checkAuthState();
    }, []);

    return (
        (() => {
            switch (authState) {
                case RouteHandlerState.LOADING:
                    return <div>Loading...</div>;
                case RouteHandlerState.NOMATCH:
                    return <Navigate to={AppRoutesNamed.AUTH} replace={true} />;
                case RouteHandlerState.MATCH:
                    return <Navigate to={AppRoutesNamed.DASHBOARD} replace={true} />;
            }
        })()
    );
}