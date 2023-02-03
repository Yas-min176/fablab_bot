function agendamento() 
{
//configuração da agenda e da planilha
    var calendarId = "your-calendar-id";
    var agenda = CalendarApp.getCalendarById(calendarId);//Setando a Agenda
    console.log(agenda)

    var planilha = SpreadsheetApp.getActiveSpreadsheet();//Setando a planilha ativa
    var sh = planilha.getActiveSheet(); //Prefixo da API da Planilha
    var row = sh.getLastRow(); //Pega o número da ultima linha com alguma informação

    var projeto = sh.getRange(row,7).getValue();//String do nome do projeto
    var quantidade = sh.getRange(row,8).getValue();//String da quantia do projeto
    var ocupacao = sh.getRange(row,2).getValue();//String da ocupação aluno/funcionário

    // @ts-ignore


    //Pegar um trecho - 
    // var now = new Date();
    // var twoHoursFromNow = new Date(now.getTime() + (18 * 60 * 60 * 1000));
    // console.log(now, twoHoursFromNow)
    // var events = agenda.getEvents(now, twoHoursFromNow);

    // console.log(events)


    // var dia_escolhido = new Date('February 17, 2022');
    // agendamentos_do_dia = agenda.getEventsForDay(dia_escolhido )

    // agendamentos_do_dia.forEach(listarAgendamentos);

    // function listarAgendamentos(val)
    // {
    //     console.log(val.getTitle())
    // }

    //var duracao = sh.getRange(row,X).getValue();


    switch(ocupacao){
        case 'Aluno':
            var MILLIS_PER_DAY = 1000 * 60 * 60 * 1;//String de duração do agendamento(atual: 1hr)
            var inicio = new Date(sh.getRange(row,5).getValue());//String de data de início do agendamento
            var fim = new Date(inicio.getTime() + MILLIS_PER_DAY);//String de data de término do agendamento
            var nome = sh.getRange(row,4).getValue();//String do nome de quem emitiu o agendamento
            var evento = quantidade + ' ' + projeto + ' - ' + nome;//String do nome do evento

            Logger.log(Utilities.formatDate(inicio,'America/Sao_Paulo', 'dd MM HH:mm:ss Z'));//Formatando a string no fuso de são paulo
            Logger.log(Utilities.formatDate(fim,'America/Sao_Paulo', 'dd MM HH:mm:ss Z'));//Formatando a string no fuso de são paulo
            console.log(fim)//Printando a string fim
            agenda.createEvent(evento, inicio, fim)//Criando o evento
        case 'Externo':
            var MILLIS_PER_DAY = 1000 * 60 * 60 * 1;//String de duração do agendamento(atual: 1hr)
            var inicio = new Date(sh.getRange(row,5).getValue());//String de data de início do agendamento
            var fim = new Date(inicio.getTime() + MILLIS_PER_DAY);//String de data de término do agendamento
            var nome = sh.getRange(row,3).getValue();//String do nome de quem emitiu o agendamento
            var evento = quantidade + ' ' + projeto + ' - ' + nome;//String do nome do evento

            Logger.log(Utilities.formatDate(inicio,'America/Sao_Paulo', 'dd MM HH:mm:ss Z'));//Formatando a string no fuso de são paulo
            Logger.log(Utilities.formatDate(fim,'America/Sao_Paulo', 'dd MM HH:mm:ss Z'));//Formatando a string no fuso de são paulo
            console.log(evento)//Printando a string fim
            agenda.createEvent(evento, inicio, fim)//Criando o evento
        case 'Funcionário':
            var MILLIS_PER_DAY = 1000 * 60 * 60 * 1;//String de duração do agendamento(atual: 1hr)
            var inicio = new Date(sh.getRange(row,5).getValue());//String de data de início do agendamento
            var fim = new Date(inicio.getTime() + MILLIS_PER_DAY);//String de data de término do agendamento
            var nome = sh.getRange(row,3).getValue();//String do nome de quem emitiu o agendamento
            var evento = quantidade + ' ' + projeto + ' - ' + nome;//String do nome do evento

            Logger.log(Utilities.formatDate(inicio,'America/Sao_Paulo', 'dd MM HH:mm:ss Z'));//Formatando a string no fuso de são paulo
            Logger.log(Utilities.formatDate(fim,'America/Sao_Paulo', 'dd MM HH:mm:ss Z'));//Formatando a string no fuso de são paulo
            console.log(evento)//Printando a string fim
            agenda.createEvent(evento, inicio, fim)//Criando o evento
    }
}