import React, { useEffect } from 'react';
import { mountTaskManager } from './taskManagerApp.js';

export default function App() {
  useEffect(() => mountTaskManager(), []);

  return (
    <>
      <div id="app" />
      <div id="modal-root" />
      <div id="toast" className="toast" />
    </>
  );
}
