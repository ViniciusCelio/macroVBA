# Macro em VBA integrando Excel e Outlook

# Objetivo
Este projeto foi criado com a finalidade de automatizar a tarefa de controlar o bloqueio e desbloqueio de usuários que estão em período de férias na empresa em que trabalho. De início, eu controlava qual data deveria desbloquear um usuário que estava voltando de férias entrando na planilha **todos os dias** e verificando se havia alguém a ser desbloqueado. Este trabalho se tornava oneroso e improdutivo e visualizando esta situação, decidi criar um botão nesta mesma planilha que pegasse os dados que estavam lá e criasse uma reunião no Outlook, sendo assim, eu precisaria apenas preencher os dados enviados pela equipe de RH mensalmente uma única vez, clicar no botão e aguardar o aviso do Outlook para realizar o desbloqueio. fdd

# Resultados
O print a seguir mostra o exemplo da planilha criada, que é possível ser baixada aqui mesmo nesse repositório (ControleDeFerias.xlsx). Os campos podem ser adaptados para outras necessidades, porém o principal para esta demanda, no meu caso, é saber o número de matrícula do colaborador, o nome do colaborador, e a data que este colaborador necessita ser desbloqueado. A coluna "Já criado?" serve para verificar se o compromisso já foi criado no Outlook e deixar marcado com um "X", a coluna "Data da criação" serve apenas para gravar a data e hora que o botão "Agendar no Outlook" foi clicado e os compromissos salvos.  

>![image](https://github.com/ViniciusCelio/macroVBA/assets/87146891/37b1b1ec-f2de-4855-821e-600c7b660ce9)

A seguir é possível ver o resultado no Outlook após rodar a macro. O evento é criado e marcado como compromisso.

>![image](https://github.com/ViniciusCelio/macroVBA/assets/87146891/c7c153ff-dd0a-49be-814f-fa85fcf21123)

15 minutos antes do horário marcado para o compromisso é enviado um alerta para avisar sobre a demanda (decidi definir 08:45:00 como um horário fixo para todos os compromissos criados). 

>![image](https://github.com/ViniciusCelio/macroVBA/assets/87146891/56402a9b-bb9c-4769-b2e7-e461ffb29bf9)

# Conclusão
Esta é uma solução bastante específica para uma situação de trabalho bastante específica, porém é algo que além de ter me agregado conhecimento em VBA foi importante para automatizar e deixar mais eficiente esta tarefa no meu trabalho. 


