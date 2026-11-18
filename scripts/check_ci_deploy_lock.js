/**
 * SAFEGUARD: BLOQUEIO DE DEPLOY MANUAL EM PRODUÇÃO (POLÍTICA DEVOPS GOVTECH)
 * Este script é acionado automaticamente pelo hook "predeploy" do Firebase CLI.
 * Ele impede que desenvolvedores realizem deploy direto de suas máquinas locais
 * para o canal de produção, exigindo a passagem pela esteira de CI/CD do GitHub Actions.
 */

const isCI = process.env.CI === "true" || process.env.GITHUB_ACTIONS === "true";
const allowManualDeploy = process.env.ALLOW_MANUAL_PROD_DEPLOY === "true";

if (!isCI && !allowManualDeploy) {
  console.error("\n" + "=".repeat(78));
  console.error("🛑 BLOQUEIO DE COMPLIANCE DEVOPS: DEPLOY MANUAL DIRETO PROIBIDO!");
  console.error("=".repeat(78));
  console.error("De acordo com as diretrizes de governança do Sistema GovTech:");
  console.error("1. Alterações em produção devem ser submetidas via Pull Request no GitHub.");
  console.error("2. A suíte de testes unitários da IN 65/2021 deve ser executada no pipeline.");
  console.error("3. A homologação deve ocorrer no canal efêmero (Preview Channel).");
  console.error("\nPara homologar uma alteração local em canal temporário:");
  console.error("   npx firebase hosting:channel:deploy preview-nome-da-sua-feature");
  console.error("\nSe você possui autorização expressa de contingência, execute:");
  console.error("   $env:ALLOW_MANUAL_PROD_DEPLOY=\"true\"; npx firebase deploy --only hosting\n");
  console.error("=".repeat(78) + "\n");
  process.exit(1);
} else {
  console.log("✅ Validação de ambiente CI/CD aprovada. Prosseguindo com o deploy.");
}
