# Contabtech — finanças pessoais

Aplicativo pessoal de controle financeiro em português do Brasil. Funciona como PWA, armazena os dados no IndexedDB do navegador e não usa backend, login, analytics ou APIs externas.

## Iniciar no computador

Requer Node.js 20.19+ (ou 22.12+).

```powershell
npm.cmd install
npm.cmd run dev
```

Abra o endereço local informado pelo Vite. Para validar a versão pronta para distribuição:

```powershell
npm.cmd test
npm.cmd run build
npm.cmd run preview
```

O servidor de desenvolvimento/preview serve os arquivos do aplicativo; ele não sincroniza nem armazena os dados financeiros.

## Publicar gratuitamente e instalar (computador e celular)

O app é um site estático: depois de publicado, você o abre por um link, instala como aplicativo e ele funciona offline. Não precisa do VS Code nem do computador ligado. Seus dados continuam só no aparelho; a hospedagem apenas entrega os arquivos do app.

### Opção A — GitHub Pages (gratuita, atualiza sozinha)

1. Crie uma conta em github.com e um repositório **privado ou público** (ex.: `contabtech`). Site no GitHub Pages gratuito exige repositório público em contas Free; o código não contém seus dados financeiros, mas não envie sua planilha nem backups para o repositório.
2. Envie a pasta do projeto (sem `node_modules`, `dist` e sem o arquivo `.xlsx`) para a branch `main`.
3. No repositório: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
4. O workflow [`.github/workflows/pages.yml`](.github/workflows/pages.yml) roda os testes, gera o build e publica. O endereço será `https://SEU-USUARIO.github.io/contabtech/`.

### Opção B — Netlify ou Cloudflare Pages (sem Git)

1. Rode `npm.cmd run build` uma vez; a pasta `dist` contém o site pronto.
2. Crie uma conta gratuita e envie a pasta `dist` pela função de upload manual (Netlify Drop / Cloudflare Pages → Direct Upload). Você recebe um link HTTPS.
3. Para atualizar o app no futuro, gere um novo `dist` e envie de novo. Os dados guardados no navegador não são afetados.

Os planos gratuitos e seus limites são definidos por cada provedor e podem mudar.

### Instalar

- **Computador (Chrome/Edge):** abra o link, clique no ícone de instalar na barra de endereço (ou menu → **Instalar Contabtech**). Cria atalho na área de trabalho/menu Iniciar e abre em janela própria.
- **Android (Chrome):** abra o link → menu ⋮ → **Instalar app**.
- **iPhone/iPad (Safari):** abra o link → **Compartilhar → Adicionar à Tela de Início**.

Instale e abra o app uma vez com internet para que ele seja guardado para uso offline. **Importante:** cada aparelho/navegador tem seus próprios dados. Para levar dados do computador para o celular, use o backup XLSX (veja abaixo) e guarde backups periodicamente; limpar os dados do navegador apaga o banco local.

## Instalação e funcionamento offline

- No computador, abra a aplicação no Chrome ou Edge e use a opção do navegador para instalar o aplicativo.
- No Android, abra o endereço HTTPS no Chrome e escolha **Instalar app** ou **Adicionar à tela inicial**. No iPhone/iPad, use **Compartilhar → Adicionar à Tela de Início** no Safari.
- A instalação PWA em celular exige que os arquivos estáticos estejam em um endereço HTTPS. O aplicativo não inclui hospedagem; uma hospedagem estática gratuita pode ser usada, mas os termos e a disponibilidade de terceiros podem mudar.
- Depois do primeiro carregamento, o service worker mantém a interface em cache para uso offline. Os dados ficam no perfil do navegador de cada dispositivo. O computador não precisa permanecer ligado depois que o aplicativo estiver instalado.
- Não há sincronização entre aparelhos. Use o backup XLSX para transferir ou recuperar dados. O backup contém informações financeiras e deve ser guardado em local privado.

## Importar a planilha

1. Abra **Dados e backup → Importar arquivo Excel**.
2. Selecione `Planilha contabilidade João.xlsx` e revise o ano detectado, a quantidade de registros e a comparação dos totais.
3. Confirme **Importar sem substituir dados existentes**.

O leitor reconhece as 12 abas mensais, os lançamentos de entrada, saídas fixas/variáveis, faturas, investimentos e dívidas/parcelas. Datas Excel são convertidas para `AAAA-MM-DD`; linhas sem dia ficam associadas ao mês da aba e podem ser corrigidas depois. Registros repetidos na mesma importação ou em importações futuras são ignorados, nunca sobrescritos. A aba `Exemplo` e linhas nomeadas como demonstração não são importadas. Pequenos valores de fatura sem vencimento e sem indicação de pagamento também são tratados como placeholders da planilha.

Também é possível importar um XLSX de movimentações exportado pelo próprio Contabtech. Ele importa entradas, saídas e investimentos; faturas, dívidas e metas são importadas a partir da planilha contábil original.

O importador compara entradas, saídas, investimentos e saldo com os resultados calculados nas abas mensais. A planilha contém uma fórmula `#REF!` em `MENU!H51`; ela não é usada como referência para os totais do aplicativo. Diferenças são mostradas antes da confirmação.

## Regras financeiras reproduzidas

| Elemento original | Regra no aplicativo |
| --- | --- |
| Entradas (mês) | Soma dos valores de entrada |
| Saídas (mês) | Soma das despesas variáveis e das despesas fixas marcadas como pagas |
| Investido (mês) | Soma dos lançamentos de investimento |
| Saldo do mês | Entradas − saídas consideradas pagas − investimentos |
| Gastos por categoria | Soma de despesas fixas e variáveis; fixas não pagas continuam visíveis neste gráfico, como na fórmula da planilha |
| Gastos por forma de pagamento | Despesas fixas/variáveis mais o valor mensal das parcelas cadastradas; faturas e dívidas informativas não entram no saldo |
| Faturas por bandeira | Valores informativos separados, sem efeito nas saídas ou no saldo |
| Comparação anual | Saldo e saídas de cada mês; metas e investimentos mensais |

As categorias e opções reconhecidas na planilha incluem: Bobeira, Alimentação, Saúde, Beleza, Roupa, Lazer, Uber/transporte, Mercado, Aluguel, Assinaturas, Essencial, Outros, Contas, Casa, Filhos, Educação, Pet, Presentes e Viagem; formas de pagamento Crédito, Débito, Dinheiro, Pix e Boleto; entradas Renda fixa e Renda variável. Opções novas identificadas em planilhas importadas também são preservadas.

## Lançamentos, exportação e restauração

- Use **＋ Lançamento** para registrar entrada, saída ou investimento. A data começa no dia atual; mês e ano são editáveis no topo. Despesas fixas têm estado de pagamento; esse estado afeta o saldo mensal. Um lançamento manual idêntico na mesma data pede confirmação antes de ser repetido.
- Navegue entre meses e anos sem alterar os dados dos demais períodos. Editar e excluir recalcula os indicadores na hora; exclusões pedem confirmação.
- Faturas e dívidas/parcelas têm telas próprias e não alteram o saldo. Parcelas aparecem na distribuição por forma de pagamento, conforme a fórmula mensal da planilha.
- Em **Dados e backup**, **Exportar movimentações XLSX** gera uma planilha Excel que pode ser importada novamente no aplicativo. **Baixar backup XLSX** salva todas as coleções e configurações para restauração completa.
- Restaurar permite mesclar dados sem substituir registros ou substituir o conjunto atual mediante confirmação explícita. Faça backup antes de substituir.

## Dados, privacidade e custos

O projeto não faz chamadas de rede para enviar registros. Os arquivos XLSX são processados localmente com ExcelJS (MIT); o Vite é usado apenas para desenvolvimento/build. Não há serviços pagos, cobrança por uso ou infraestrutura obrigatória no aplicativo. O armazenamento do navegador não é uma cópia de segurança e não é sincronizado; dados podem ser perdidos se o perfil/navegador for removido, então mantenha backups periódicos. IndexedDB não substitui criptografia do dispositivo: proteja o acesso ao seu usuário e ao arquivo de backup.

Hospedagem HTTPS para disponibilizar a instalação em outro aparelho é opcional e não faz parte deste repositório. O custo e a continuidade de qualquer hospedagem externa dependem do provedor escolhido; o aplicativo continua sem serviço de nuvem e sem sincronização.
