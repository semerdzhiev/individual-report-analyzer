# Individual Report Analyzer

## Какво има тук?

`analyze.ps1` е прост PowerShell скрипт, който обработва индивидуални отчети подадени за някакъв период от време, сумира заетостта (аудиторна и обща) и извежда отчет с окончателните суми.

## Общи бележки

Преди да пуснете скрипта, съберете всички индивидуални отчети, които искате да обработите, в една обща директория. Тази директория се подава на скрипта, който след това намира в нея всички файлове с разширение `.xlsx`, обработва ги и извежда получените резултати.

Няма значение как отчетите ще бъдат организирани в директорията. Например, може да са подредени в отделни директории за всеки семестър или всички файлове да са натрупани в една обща директория. Скриптът претърсва рекурсивно директорията и намира всички `.xlsx` файлове, независимо къде се намират. Това, което е важно е:

1. Скриптът предполага, че файловете са от един отчетен период, който искате да анализирате. Съответно, за всеки човек, за който има един или повече отчетни файлове в директорията, се сумират часовете от всички отчети. Така, ако в директорията са събрани отчети само за един семестър, ще се видят стойностите само за този семестър. Ако са събрани отчети за една година, ще се види годишният норматив за всеки човек. Ако са събрани отчети за две години, ще се види общият брой на часовете за тези години и т.н.

2. Скриптът не прави опит да разпознае дали файловете са валидни отчетни файлове или не. Затова в директорията не трябва да има други `.xlsx` файлове, освен файлове с отчети, които да се обработят.

## Какво е нужно за пускане на скрипта?

Скриптът създава и използва Excel COM обект, за да зареди и обработи файловете с отчетите. Затова на компютъра, на който ще пуснете скрипта, трябва да имате инсталиран Excel.

# Пускане на скрипта

Нека допуснем, че отчетите се намират в директорията `C:\Temp\Individual reports`.

Скриптът може да се пусне например така:

```powershell
.\analyze.ps1 "C:\Temp\Individual reports"
```

или ако искате да изпишете пълната форма:

```powershell
.\analyze.ps1 -InputDirectory "C:\Temp\Individual reports"
```

Скриптът ще създаде Excel COM обект и ще го използва за обработката на отчетите. **Моля, не затваряйте прозореца на Excel, който ще се появи, докато върви обработката на отчетите.**

Когато скриптът приключи, в прозореца на Excel ще видите получения резултат.

Ако искате да запишете резултатите във файл, можете да го направите ръчно през Excel или да подадете съответния параметър на скрипта:

```powershell
.\analyze.ps1 .\reports\ -CsvOutput result.csv
```

Ако искате в отчета да фигурират само крайните суми, без стойностите извлечени от отделните отчети, добавете параметъра `-TotalsOnly`. Например:

```powershell
.\analyze.ps1 .\reports\ -TotalsOnly
```

По подразбиране скриптът сравнява сумите получени за всеки човек спрямо годишния норматив за един преподавал на пълен щат (270 часа аудиторна и 360 часа обща заетост). Ако искате да промените тези стойонсти, можете да го направите през параметрите `-AuditoryHoursNorm` и `-TotalHoursNorm`. Например, за да зададем стойности като за половин щат, можем да използваме:

```powershell
.\analyze.ps1 .\reports\ -AuditoryHoursNorm 135 -TotalHoursNorm 180
```