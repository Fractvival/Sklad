# Sklad
###### Sklad techniků Siemens



### Instalace
<hr />

- Nejprve je nutné stáhnout samotný program Skladu [ZDE](https://github.com/Fractvival/Sklad/blob/master/!Sklad.xlsm).
   (Program umísti tam, odkud jej plánuješ spouštět)
- Před spuštěním je potřeba vytvořit pracovní složku skladu, ve které bude uložen jak sešit s díly ve skladu, tak veškeré zálohy a informace o uživatelích.
 (typické místo pro umístění složky je v kořenu disku, například: `C:\Sklad`
 
 *Samozřejmě, že by program mohl složku při prvním spuštění vytvořit sám, jde ale o to, že se očekává složka skladu umístěná v sítí. V takovém případě, při nedostupnosti sítě (tedy i složky), se uživateli díky tomu zobrazí hlášení o nedostupnosti, a teprve poté si může zvolit jiné místo. Přesně toto se bude dít i při prvním spuštění.*
 
 - Do nově vytvořené složky nakopíruj soubor se seznamem dílů ve skladě. V tomto souboru se očekávají jistá pravidla jeho struktury jak bude popsáno dále. Dále se očekává, že se tento soubor bude jmenovat `Sklad.xlsx`, ovšem názvy a umístění lze později změnit podle vlastní libovůle.
 
 Následující tabulka struktury souboru s náhradními díly:
 
|      KZM      |  Part NUMBER  |    Nazev1     |    Nazev2     |     Počet     | Inventura SKLADU |   Umístění    | Doplněno DNE  |
| ------------- | ------------- | ------------- | ------------- | ------------- | ---------------- | ------------- | ------------- |
|   790040409   |22.1229.043-01 |   SPR.HELIC   |               |       1       | **NEPOUZIVA SE!**    |      3C1      | **NEPOUZIVA SE**  |
|   790040397   |25.1033.206-04 |     PULLEY    |      GR       |       45      | **NEPOUUIVA SE!**    |      2C17     | **NEPOUZIVA SE**  |
 
Sloupce `Inventura SKLADU` (F) a `Doplněno DNE` (H) se aktuálně nepoužívají, avšak se s nimi počítá pro možnost jejich využití pro jiné účely.

### Spuštění
<hr />

Jakmile je instalace hotová, může se program spustit. Po spuštění se tedy program zeptá uživatele, jestli chce zvolit složku se skladem. Jakmile je složka zvolená, provede se její inicializace- to znamená, že se v ní vytvoří podsložky pro logy uživatelů a pro zálohy. Teprve poté je již zobrazen úvodní dialog pro přihlášení uživatele.

Při prvním spuštění je automaticky vytvořen jediný uživatel, a to Administrátor.<br /> Přihlašovací heslo je: `123456`

Po přihlášení, v nabídce `Správce`, je poté doporučeno spustit `Editor uživatelů`, vytvořit nového správce, odhlásit se a při dalším přihlášení, jako nový správce, toho defaultního odstranit. Pak už nastává editace všech ostaních uživatelů.
