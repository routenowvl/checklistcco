// Teste da função getRouteDateForCurrentTime - Regra D+1 após 21:00h
// Regra:
// - 21:00h às 23:59h: Retorna data de amanhã (D+1)
// - 00:00h às 20:59h: Retorna data de hoje (D)

import { getRouteDateForCurrentTime, getBrazilDate } from './utils/dateUtils';

console.log('=== Teste getRouteDateForCurrentTime ===\n');

// Testa o comportamento atual (baseado no horário de Brasília agora)
const resultado = getRouteDateForCurrentTime();
const hoje = getBrazilDate();

// Calcula amanhã para comparação
const tomorrow = new Date();
tomorrow.setDate(tomorrow.getDate() + 1);
const amanha = tomorrow.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' })
  .split('/')
  .reverse()
  .join('-');

console.log(`📅 Data de hoje (Brasília): ${hoje}`);
console.log(`📅 Data de amanhã (Brasília): ${amanha}`);
console.log(`📅 Resultado da função: ${resultado}`);

// Verifica se o resultado é coerente com o horário atual
const now = new Date();
const hoursBr = now.getHours(); // Hora local (considerando que o sistema está em UTC-3)

console.log(`\n🕐 Horário atual (local): ${hoursBr}:00`);

if (hoursBr >= 21) {
  console.log(`\n✅ Regra aplicada: Horário >= 21:00h → Deve usar D+1`);
  if (resultado === amanha) {
    console.log(`✅ TESTE PASSOU: Resultado correto (${resultado} === ${amanha})`);
  } else {
    console.log(`❌ TESTE FALHOU: Esperado ${amanha}, obtido ${resultado}`);
    process.exit(1);
  }
} else {
  console.log(`\n✅ Regra aplicada: Horário < 21:00h → Deve usar D (hoje)`);
  if (resultado === hoje) {
    console.log(`✅ TESTE PASSOU: Resultado correto (${resultado} === ${hoje})`);
  } else {
    console.log(`❌ TESTE FALHOU: Esperado ${hoje}, obtido ${resultado}`);
    process.exit(1);
  }
}

console.log('\n🎉 Teste concluído com sucesso!');
