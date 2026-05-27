/**
 * Arredonda um valor monetário para a dezena de centavos (1 casa decimal)
 * usando a regra matemática padrão (round half up).
 *
 * Exemplos:
 *   9.23  -> 9.2  (4 e abaixo desce)
 *   9.25  -> 9.3  (5 e acima sobe)
 *   9.27  -> 9.3
 *   9.04  -> 9.0
 *   9.05  -> 9.1
 *
 * Isso é aplicado a TODOS os valores monetários que entram no sistema,
 * seja por importação de planilha ou digitação manual.
 */
export function roundMoney(value: number | null | undefined): number {
  if (value == null || !Number.isFinite(value)) {
    return 0;
  }

  // Math.round em JS arredonda meio para cima (5 sobe).
  // Multiplica por 10 -> arredonda -> divide por 10 = 1 casa decimal.
  return Math.round(value * 10) / 10;
}
