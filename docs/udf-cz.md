# Funkce

Balíèek rozšíøuje vzorce (funkce) Excelu o vybrané statistické funkce, které urychlují výpoèty metod vyuèovaných v pøedmìtech ST201 a ST210. Funkce obsahují základní dokumentaci výstupù a vstupních parametrù.

Doporuèuji zaèít funkce zaèít používat buïto tehdy, když budete mít dobøe zvládnuté ruèní výstupy, pøípadnì jako kontrolu nad ruènì vypoèítaným výsledkem.

**Popisné statistiky**

| **název funkce** | **popis funkce** |
| :--- | :--- |
| [MEAN.W](#Vážený-aritmetický-prùmìr) | vážený aritmetický prùmìr |
| [HARMEAN.W](#Vážený-harmonický-prùmìr) | vážený harmonický prùmìr |
| [GEOMEAN.W](#Vážený-geometrický-prùmìr) | vážený geometrický prùmìr |
| [VAR.P.W](#Vážený-populaèní-rozptyl)|	vážený populaèní rozptyl |
| [VAR.S.W](#Vážený-rozptyl-ze-vzorku) | vážený rozptyl ze vzorku |
| [SMODCH.P.W](#Vážená-smìrodatná-odchylka-populace) | vážená smìrodatná odchylka populace |
| [SMODCH.S.W](#Vážená-smìrodatná-odchylka-ze-vzorku) | vážená smìrodatná odchylka ze vzorku |
| [VAR.RANGE](#Variaèní-rozpìtí) | variaèní rozpìtí souboru |
| [ROZKLAD.ROZPTYLU](#Rozklad-rozptylu) | rozklad rozptylu |
| [MAD](#Absolutní-mediánová-odchylka) | absolutní mediánová odchylka |

**Distribuèní funkce**

| **název funkce** | **popis funkce** |
| :--- | :--- |
| [NORM.DIST.RANGE](#Interval-normálního-rozdìlení) | pravdìpodobnost intervalu u normálního rozdìlení |

**Analýza kontingenèní tabulky**

| **název funkce** | **popis funkce** |
| :--- | :--- |
| [KONTINGENCE.G](#Testová-statistika-G) | Testová statistika G |
| [KONTINGENCE.PV](#P-hodnota-pro-kontingenci) | P-hodnota pro kontingenci |
| [KONTINGENCE.C](#Pearsonùv-koeficient-kontingence-C) | Pearsonùv koeficient kontingence C |
| [KONTINGENCE.V](#Cramérùv-koeficient-kontingence-V) | Cramérùv koeficient kontingence V |

**Spearmanùv korelaèní koeficient**
| **název funkce** | **popis funkce** |
| :--- | :--- |
| [SPEARMAN](#Spearmanùv-korelaèní-koeficient) | Spearmanùv korelaèní koeficient |
| [SPEARMAN.PV](#P-hodnota-pro-Spearmanùv-test) | P-hodnota pro Spearmanovo rho |
| [SPEARMAN.T](#Testová-statistika-T-pro-Spermanùv-test) | Pearsonùv koeficient kontingence C |





## Popisné statistiky

### Vážený aritmetický prùmìr

```Excel
MEAN.W
```

Vypoèítá vážený aritmetický prùmìr. Oblast s hodnotami musí mít stejný rozsah (stejný poèet bunìk) jako oblast s váhami. Zadávejte bez záhlaví (pouze èísla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami, které mají být prùmìrovány |
| **váhy** | oblast bunìk s váhami |

 ```Excel
 =PRÙMÌR.W(A2:A10;B2:B10)
 ```

### Vážený harmonický prùmìr

```Excel
HARMEAN.W
```

Vypoèítá vážený harmonický prùmìr. Oblast s hodnotami musí mít stejný rozsah (stejný poèet bunìk) jako oblast s váhami. Zadávejte bez záhlaví (pouze èísla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami, které mají být prùmìrovány |
| **váhy** | oblast bunìk s váhami |

 ```Excel
 =HARMEAN.W(A2:A10;B2:B10)
 ```

### Vážený geometrický prùmìr

```Excel
GEOMEAN.W
```

Vypoèítá vážený geometrický prùmìr. Oblast s hodnotami musí mít stejný rozsah (stejný poèet bunìk) jako oblast s váhami. Zadávejte bez záhlaví (pouze èísla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami, které mají být prùmìrovány |
| **váhy** | oblast bunìk s váhami |

 ```Excel
 =GEOMEAN.W(A2:A10;B2:B10)
 ```

### Vážený populaèní rozptyl

```Excel
VAR.P.W
```

Spoèítá vážený rozptyl pro populaci.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s dílèími prùmìry |
| **váhy** | oblast bunìk s váhami |

```Excel
 =VAR.P.W(A2:A10;B2:B10)
```

### Vážený rozptyl ze vzorku

```Excel
VAR.S.W
```

Spoèítá vážený rozptyl pro vzorek.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s dílèími prùmìry |
| **váhy** | oblast bunìk s váhami |

```Excel
=VAR.S.W(A2:A10;B2:B10)
```

### Vážená smìrodatná odchylka populace

```Excel
SMODCH.P.W
```

Spoèítá váženou smìrodatnou odchylku pro populaci.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s dílèími prùmìry |
| **váhy** | oblast bunìk s váhami |

```Excel
=SMODCH.P.W(A2:A10;B2:B10)
```

### Vážená smìrodatná odchylka ze vzorku

```Excel
SMODCH.S.W
```

Spoèítá váženou smìrodatnou odchylku pro vzorek.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s dílèími prùmìry |
| **váhy** | oblast bunìk s váhami |

```Excel
=SMODCH.S.W(A2:A10;B2:B10)
```

### Variaèní rozpìtí

```Excel
VAR.RANGE
```

Spoèítá variaèní rozpìtí souboru, tedy rozdíl mezi nejvìtší a nejmìnší hodnotou v souboru.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami |

```Excel
=VAR.RANGE(A2:A10)
```

### Rozklad rozptylu

```Excel
ROZKLAD.ROZPTYLU
```

Spoèítá vìtu o rozkladu rozptylu. Pro výpoèet je nutné zadat dílèí rozptyly, støední hodnoty a váhy.

**argumenty**

| id | popis |
| --- | --- |
| **rozptyly** | oblast bunìk s dílèími rozptyly |
| **prùmìry** | oblast bunìk s dílèími prùmìry |
| **váhy** | oblast bunìk s váhami |

```Excel
=ROZKLAD.ROZPTYLU(A2:A10;B2:B10;C2:C10)
```

### Absolutní mediánová odchylka

```Excel
MAD
```

Spoèítá absolutní mediánovou odchylku souboru.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami |

```Excel
=MAD(A2:A10)
```

## Distribuèní funkce

### Interval normálního rozdìlení

```Excel
NORM.DIST.RANGE
```

Spoèítá pravdìpodobnost jevu mezi dvìma referenèními body u velièiny s normálním rozdìlením. Funguje podobnì jako funkce BINOM.DIST.RANGE.

**argumenty**

| id | popis |
| --- | --- |
| **x** | støední hodnota rozdìlení |
| **s** | smìrodatná odchylka rozdìlení |
| **x1** | spodní hranice |
| **x2** | horní hranice |

```Excel
=NORM.DIST.RANGE(10;5;6;8)
```

## Analýza kontingenèní tabulky

### Testová statistika G

```Excel
KONTINGENCE.G
```

Vypoète testovou statistiku G z kontingenèní (køížové) tabulky. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.G(B2:D5)
 ```

### P-hodnota pro kontingenci

```Excel
KONTINGENCE.PV
```

Vypoète p-hodnotu pro kontingenèní (køížovou) tabulku. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.PV(B2:D5)
 ```

### Pearsonùv koeficient kontingence C

```Excel
KONTINGENCE.C
```

Vypoète testovou statistiku Pearsonova C z kontingenèní (køížové) tabulky. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.C(B2:D5)
 ```

### Cramérùv koeficient kontingence V

```Excel
KONTINGENCE.V
```

Vypoète testovou statistiku Cramérovo V z kontingenèní (køížové) tabulky. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.V(B2:D5)
 ```

## Spearmanùv korelaèní koeficient a testy

```Excel
SPEARMAN;
```

### Spearmanùv korelaèní koeficient

Vypoèítá Spearmanùv korelaèní koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bunìk s hodnotami øady x |
| **hodnoty y** | oblast bunìk s hodnotami øady y |

```Excel
=SPEARMAN(A2:A21;B2:B21)
```

### Testová statistika T pro Spermanùv test

```Excel
SPEARMAN.T
```

Vypoèítá testovou statistiku T pro Spearmanùv korelaèní koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bunìk s hodnotami øady x |
| **hodnoty y** | oblast bunìk s hodnotami øady y |

```Excel
=SPEARMAN.T(A2:A21;B2:B21)
```

### P-hodnota pro Spearmanùv test

```Excel
SPEARMAN.PV
```

Vypoèítá p-hodnotu pro Spearmanùv korelaèní koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bunìk s hodnotami øady x |
| **hodnoty y** | oblast bunìk s hodnotami øady y |

```Excel
=SPEARMAN.PV(A2:A21;B2:B21)
```