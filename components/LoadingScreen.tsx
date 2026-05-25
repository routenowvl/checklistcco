import React from 'react';

const LoadingScreen: React.FC = () => {
  return (
    <div className="splash-loading">
      <div className="splash-bg-glow" />
      <div className="splash-inner">
        <img src="/logo.png" alt="VIA" className="splash-logo-anim" />
        <div className="splash-dots-anim">
          <span></span><span></span><span></span>
        </div>
      </div>
      <style>{`
        .splash-loading {
          position: fixed; inset: 0; z-index: 99999;
          display: flex; align-items: center; justify-content: center;
          background: #05080f;
        }
        .splash-bg-glow {
          position: absolute; inset: 0;
          background: radial-gradient(ellipse at 50% 50%, rgba(0,212,255,0.08) 0%, transparent 70%);
        }
        .splash-inner {
          position: relative; z-index: 1;
          display: flex; flex-direction: column; align-items: center; gap: 32px;
        }
        .splash-logo-anim {
          width: 220px; height: auto;
          animation: splashFloat 3s ease-in-out infinite;
          filter: drop-shadow(0 0 40px rgba(0,212,255,0.25));
        }
        @keyframes splashFloat {
          0%, 100% { transform: translateY(0px); }
          50% { transform: translateY(-18px); }
        }
        .splash-dots-anim {
          display: flex; gap: 10px;
        }
        .splash-dots-anim span {
          width: 8px; height: 8px; border-radius: 50%;
          background: rgba(0,212,255,0.4);
          animation: splashPulse 1.4s ease-in-out infinite;
        }
        .splash-dots-anim span:nth-child(2) { animation-delay: 0.2s; }
        .splash-dots-anim span:nth-child(3) { animation-delay: 0.4s; }
        @keyframes splashPulse {
          0%, 80%, 100% { transform: scale(0.6); opacity: 0.3; background: rgba(0,212,255,0.3); }
          40% { transform: scale(1.2); opacity: 1; background: rgba(0,212,255,0.9); }
        }
      `}</style>
    </div>
  );
};

export default LoadingScreen;
