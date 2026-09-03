import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import ErrorBoundary from './components/ErrorBoundary.tsx';
import './index.css';

// Suppress harmless Firebase network connection warnings in the console
const originalConsoleError = console.error;
console.error = (...args) => {
  if (
    typeof args[0] === 'string' &&
    (args[0].includes('Could not reach Cloud Firestore backend') ||
      args[0].includes('code=unavailable') ||
      args[0].includes('auth/network-request-failed'))
  ) {
    return;
  }
  originalConsoleError(...args);
};

const originalConsoleWarn = console.warn;
console.warn = (...args) => {
  if (
    typeof args[0] === 'string' &&
    (args[0].includes('Could not reach Cloud Firestore backend') ||
      args[0].includes('code=unavailable') ||
      args[0].includes('auth/network-request-failed'))
  ) {
    return;
  }
  originalConsoleWarn(...args);
};

window.addEventListener('unhandledrejection', (event) => {
  if (
    event.reason &&
    (event.reason.code === 'auth/network-request-failed' ||
      (typeof event.reason.message === 'string' &&
        event.reason.message.includes('auth/network-request-failed')))
  ) {
    event.preventDefault();
  }
});

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
);
