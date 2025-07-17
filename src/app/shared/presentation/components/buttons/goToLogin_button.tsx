import { AppRoutesNamedFullPath } from "@routes/app_routes_named";
import SecondaryButton from "./Secondary_button";
import { useNavigate } from "react-router-dom";

export default function GoToLoginButton() {
  const navigate = useNavigate();

  const handleOnLogin = () => {
    navigate(AppRoutesNamedFullPath.SIGNIN, { replace: true });
  };

  return <SecondaryButton label="Go to login" onPress={handleOnLogin} />;
}
