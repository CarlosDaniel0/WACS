# Windows Auto Clean Spreadsheets

Serviço do Microsoft Windows que executa a limpeza de planilhas do excel (.xlsx)
após o sistema iniciar.

## Início Rápido

É necessário possuir o .NET 7 instalado para executar esse projeto

executar
```
dotnet run .
```

Crie seu Serviço do Windows [aqui](https://learn.microsoft.com/pt-br/dotnet/core/extensions/windows-service)

## Pacotes

* [ClosedXML](https://www.nuget.org/packages/ClosedXML/)
* [Miscrosoft.Extensions.Hosting](https://www.nuget.org/packages/Microsoft.Extensions.Hosting)
* [Microsoft.Extensions.Hosting.WindowsServices](https://www.nuget.org/packages/Microsoft.Extensions.Hosting.WindowsServices)
* [Nager.Holiday](https://www.nuget.org/packages/Nager.Holiday)
* [Quartz.AspNetCore](https://www.nuget.org/packages/Quartz.AspNetCore)