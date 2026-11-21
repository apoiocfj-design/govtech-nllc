/**
 * DICIONÁRIO EDITORIAL E CONCORDÂNCIA FORMAL (PT-BR CULTO PARA AUDITORIA TCU/CGU/AGU)
 * Elimina abreviações informais ("p/ Anexar") e falsos plurais ("cotação(ões) válida(s)").
 */

import { formatBrlCurrency, getActiveLegalThreshold } from "../../domain/config/legalThresholds.ts";

export function pluralizeValidQuotes(count: number): string {
  const n = Math.max(0, Math.trunc(Number(count) || 0));
  if (n === 1) return "1 cotação válida";
  return `${n} cotações válidas`;
}

export function pluralizeObtainedQuotes(count: number): string {
  const n = Math.max(0, Math.trunc(Number(count) || 0));
  if (n === 1) return "1 cotação obtida";
  return `${n} cotações obtidas`;
}

export function pluralizeItems(count: number): string {
  const n = Math.max(0, Math.trunc(Number(count) || 0));
  if (n === 1) return "1 item";
  return `${n} itens`;
}

export const PT_BR_EDITORIAL = {
  autoJustificationHeader:
    "Justificativa Automática para Anexar ao Processo Administrativo (Art. 6º da IN SEGES/ME nº 65/2021):",
  nonFractionationHeading:
    "Classificação para Controle de Não Fracionamento (Art. 75, § 1º da Lei nº 14.133/2021)",
  restrictedBasketAlertTitle(validCount: number): string {
    return `Auditoria: Cesta com ${pluralizeValidQuotes(validCount)}`;
  },
  restrictedBasketJustification(validCount: number): string {
    const n = Math.max(0, Math.trunc(Number(validCount) || 0));
    if (n >= 3) {
      return `Atendido o requisito mínimo legal de 3 (três) preços válidos (${pluralizeValidQuotes(n)}), conforme art. 6º, § 5º da IN SEGES/ME nº 65/2021.`;
    }
    const verb = n === 1 ? "Obteve-se apenas 1 cotação válida" : `Obtiveram-se apenas ${pluralizeValidQuotes(n)}`;
    return `ATENÇÃO AUDITORIA: ${verb}, exigindo justificativa circunstanciada e aprovação expressa da autoridade competente (art. 6º, § 5º da IN SEGES/ME nº 65/2021).`;
  },
  fracionamentoBannerDescription(pdmOrClass: string, accumulatedValue: number, fiscalYear: number = 2026): string {
    const cfg = getActiveLegalThreshold(fiscalYear);
    return (
      `O valor acumulado neste lote para o PDM/Classe "${pdmOrClass}" totaliza ${formatBrlCurrency(accumulatedValue)}. ` +
      `Caso a UASG possua outras contratações diretas deste mesmo ramo de atividade no exercício corrente que somadas ` +
      `ultrapassem o teto legal de ${formatBrlCurrency(cfg.art75IncisoII)} (${cfg.decreeNumber}), a dispensa será vedada ` +
      `pelo TCU, tornando-se obrigatória a instauração de Pregão Eletrônico.`
    );
  },
} as const;
