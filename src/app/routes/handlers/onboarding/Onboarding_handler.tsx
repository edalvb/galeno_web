import { useEffect, useState } from "react";
import { AppRoutesNamed } from "../../app_routes_named";
import { Navigate } from "react-router-dom";
import { RouteHandlerState } from "@shared/domain/router_handler_state";
import { AuthService } from "@shared/application/auth_service";
import OnboardingRootView from "@features/onboarding/presentation/root/Onboarding_root_view";

export default function OnboardingHandler() {

    let [authState, setAuthState] = useState(RouteHandlerState.LOADING);

    useEffect(() => {
        const checkAuthState = async () => {
            const state = await AuthService.getAuthState();
            setAuthState(state ? RouteHandlerState.NOMATCH : RouteHandlerState.MATCH);
        };

        checkAuthState();
    }, [])

    return (
        (() => {
            switch (authState) {
                case RouteHandlerState.LOADING:
                    return <div>Loading...</div>;
                case RouteHandlerState.NOMATCH:
                    return <Navigate to={AppRoutesNamed.DASHBOARD} replace={true} />;
                case RouteHandlerState.MATCH:
                    return <OnboardingRootView />;
            }
        })()
    );
}