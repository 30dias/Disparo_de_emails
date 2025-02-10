-=-= Alguns detalhes sobre o uso do Script de envio de emails =-=-

-As planilhas que serão utilizadas como fonte devem permanecer no mesmo diretório que o executável. Isso acontece pois o código utiliza um caminho fixo para encontrar o arquivo na máquina.
-Quando o executável "script" iniciar, clique sobre o terminal de comando, pressione ENTER e aguarde alguns instantes.

-A leitura e filtragem das planilhas é totalmente automatizada, no entanto, para realizar o disparo do email você precisa de uma conta do GMAIL. Mais informações sobre isso:
Na primeira execução dos disparos de email, o terminal irá dar algumas instruções breves sobre como e porque utilizar um GMAIL.

O servidor do gmail é público e de fácil acesso, portanto escolhi ele para realizar o projeto demonstrativo. Caso seja necessário utilizar um domínio de email privado basta mudar o servidor SMTP e a porta específica.

--IMPORTANTE--
Sobre a senha.
O GMAIL não permite utilizar a senha da conta diretamente em aplicativos externos, por isso, para o disparo dos emails é necessário obter uma senha de app da Google. Para fazer isso alguns passos devem ser seguidos:

1-É necessário que a conta tenha a autenticação de 2 fatores ativada;
2-Acesse o site https://myaccount.google.com/apppasswords;
3-Gere uma nova senha de app digitando no campo o nome que preferir;
4-Copie a senha e cole no terminal de comando no momento em que ele pede.

Só será absolutamente necessário realizar esse procedimento na primeira execução do script, as próximas usarão um bloco de notas para realizar a leitura do login.
*OBS: A senha de app é segura desde que permaneça bem guardada. Você pode desativa-la a qualquer momento na mesma página em que a criou.

Detalhes adicionais:
O script atual foi desenvolvido com o propósito de demonstrar a automação em andamento.
pontos de melhoria para o projeto final: formatação responsiva do HTML, adição dos links oficiais no corpo do email, adição das imagens e logos.