import { createCanvas } from 'canvas';
import { writeFileSync } from 'fs';
import { join } from 'path';

const publicDir = join(process.cwd(), 'public');

function drawIcon(size) {
  const canvas = createCanvas(size, size);
  const ctx = canvas.getContext('2d');
  
  // Background gradient
  const gradient = ctx.createLinearGradient(0, 0, size, size);
  gradient.addColorStop(0, '#2563eb');
  gradient.addColorStop(1, '#1d4ed8');
  
  // Rounded rectangle background
  const radius = size * 0.195;
  ctx.beginPath();
  ctx.roundRect(0, 0, size, size, radius);
  ctx.fillStyle = gradient;
  ctx.fill();
  
  // White square with checkmark
  const squareSize = size * 0.53;
  const squareX = (size - squareSize) / 2;
  const squareY = (size - squareSize) / 2 - size * 0.05;
  const squareRadius = size * 0.078;
  
  ctx.beginPath();
  ctx.roundRect(squareX, squareY, squareSize, squareSize, squareRadius);
  ctx.fillStyle = 'white';
  ctx.fill();
  
  // Checkmark
  ctx.beginPath();
  ctx.lineWidth = size * 0.068;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = '#2563eb';
  
  const checkStartX = squareX + squareSize * 0.3;
  const checkStartY = squareY + squareSize * 0.55;
  const checkMidX = squareX + squareSize * 0.45;
  const checkMidY = squareY + squareSize * 0.72;
  const checkEndX = squareX + squareSize * 0.75;
  const checkEndY = squareY + squareSize * 0.35;
  
  ctx.moveTo(checkStartX, checkStartY);
  ctx.lineTo(checkMidX, checkMidY);
  ctx.lineTo(checkEndX, checkEndY);
  ctx.stroke();
  
  // Text "CCO"
  ctx.fillStyle = 'white';
  ctx.font = `bold ${size * 0.156}px Arial`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('CCO', size / 2, size * 0.88);
  
  return canvas;
}

// Generate icons
const sizes = [
  { size: 192, filename: 'pwa-192x192.png' },
  { size: 512, filename: 'pwa-512x512.png' },
  { size: 180, filename: 'apple-touch-icon.png' },
  { size: 32, filename: 'favicon.png' }
];

sizes.forEach(({ size, filename }) => {
  const canvas = drawIcon(size);
  const buffer = canvas.toBuffer('image/png');
  writeFileSync(join(publicDir, filename), buffer);
  console.log(`✓ Generated ${filename} (${size}x${size})`);
});

console.log('\nÍcones gerados com sucesso na pasta public/');
