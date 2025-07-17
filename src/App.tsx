import { useEffect, useState } from "react";
import { AppConfig } from "@config/app_config";
import AppRouter from "./app/routes/App_router";
import { spinnerStateProvider } from "@shared/presentation/components/loading/spinner_state";
import CustomSpinner from "@shared/presentation/components/loading/Spinner";

function App() {
  const [isLoading, setIsLoading] = useState(true);
  const showSpinner = spinnerStateProvider((state) => state.showSpinner);

  useEffect(() => {
    const initializeApp = async () => {
      await AppConfig.initialize();
      setIsLoading(false);
    };
    initializeApp();
  }, []);

  return isLoading ? (
    <div>Loading...</div>
  ) : (
    <div className="primaryMainTheme text-foreground bg-background">
      <AppRouter />
      {showSpinner && <CustomSpinner />}
    </div>
  );
}

export default App;
