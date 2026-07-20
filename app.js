function adicionarCalendario(titulo,data,hora){

let evento =
`
BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
SUMMARY:${titulo}
DTSTART:${data}T${hora}
END:VEVENT
END:VCALENDAR
`;

let arquivo = new Blob([evento],{
    type:"text/calendar"
});

let link=document.createElement("a");

link.href=URL.createObjectURL(arquivo);

link.download="evento.ics";

link.click();

}
