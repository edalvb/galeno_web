import "reflect-metadata";
import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import { HeroUIProvider } from "@heroui/react";
import { Toaster } from 'react-hot-toast';

import 'react-phone-number-input/style.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <HeroUIProvider locale="en-US">
      <Toaster position="bottom-right" containerClassName="text-sm" />
      <App />
    </HeroUIProvider>
  </StrictMode>,
)
