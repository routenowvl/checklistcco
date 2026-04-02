import React from 'react';

const LoadingScreen: React.FC = () => {
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-amber-50 to-slate-100 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950 flex flex-col items-center justify-center p-6">
      {/* Container da animação */}
      <div className="relative flex items-center justify-center gap-1">
        {/* Barras ondulando */}
        {[...Array(5)].map((_, i) => (
          <div
            key={`bar-${i}`}
            className="w-2 h-12 rounded-full bg-white/80 dark:bg-slate-300/80 animate-wave"
            style={{ 
              animationDelay: `${i * 0.12}s`,
              animationDuration: '1s'
            }}
          />
        ))}
      </div>

      {/* Estilos de animação */}
      <style>{`
        @keyframes wave {
          0%, 100% {
            transform: scaleY(0.4);
            opacity: 0.3;
          }
          50% {
            transform: scaleY(1);
            opacity: 1;
          }
        }

        .animate-wave {
          animation: wave 1s ease-in-out infinite;
        }
      `}</style>
    </div>
  );
};

export default LoadingScreen;
