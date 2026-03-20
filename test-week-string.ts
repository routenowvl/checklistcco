// Teste da função getWeekString - Lógica do Excel
// =SE(C3="";"";MAIÚSCULA(TEXTO(C3;"mmm"))&" S"&(1+(DIA(C3)>7)+(DIA(C3)>15)+(DIA(C3)>22)))

import { getWeekString } from './utils/dateUtils';

console.log('=== Teste getWeekString ===\n');

// Testes com diferentes datas
const testData = [
  { date: '2026-01-05', expected: 'JAN S1' },   // Dia 5 → S1
  { date: '2026-01-10', expected: 'JAN S2' },   // Dia 10 → S2
  { date: '2026-01-18', expected: 'JAN S3' },   // Dia 18 → S3
  { date: '2026-01-25', expected: 'JAN S4' },   // Dia 25 → S4
  { date: '2026-02-03', expected: 'FEV S1' },   // Dia 3 → S1
  { date: '2026-02-14', expected: 'FEV S2' },   // Dia 14 → S2
  { date: '2026-02-20', expected: 'FEV S3' },   // Dia 20 → S3
  { date: '2026-02-28', expected: 'FEV S4' },   // Dia 28 → S4
  { date: '2026-03-01', expected: 'MAR S1' },   // Dia 1 → S1
  { date: '2026-03-08', expected: 'MAR S2' },   // Dia 8 → S2
  { date: '2026-03-16', expected: 'MAR S3' },   // Dia 16 → S3
  { date: '2026-03-23', expected: 'MAR S4' },   // Dia 23 → S4
  { date: '2026-12-31', expected: 'DEZ S4' },   // Dia 31 → S4
  { date: '05/01/2026', expected: 'JAN S1' },   // Formato BR
  { date: '15/02/2026', expected: 'FEV S2' },   // Formato BR
  { date: '', expected: '' },                    // String vazia
  { date: 'invalido', expected: '' },            // Data inválida
];

let passed = 0;
let failed = 0;

testData.forEach(({ date, expected }) => {
  const result = getWeekString(date);
  const status = result === expected ? '✅' : '❌';
  
  if (result === expected) {
    passed++;
  } else {
    failed++;
  }
  
  console.log(`${status} getWeekString('${date}') = '${result}' (esperado: '${expected}')`);
});

console.log(`\n=== Resumo ===`);
console.log(`✅ Passaram: ${passed}`);
console.log(`❌ Falharam: ${failed}`);
console.log(`Total: ${testData.length}`);

if (failed === 0) {
  console.log('\n🎉 Todos os testes passaram!');
} else {
  console.log('\n⚠️ Alguns testes falharam!');
  process.exit(1);
}
