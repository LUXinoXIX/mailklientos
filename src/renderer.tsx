import { createRoot } from 'react-dom/client';
import { App } from './renderer/App';
import { LockGate } from './renderer/components/LockScreen';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <>
    <App />
    <LockGate />
  </>,
);
