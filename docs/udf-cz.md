# Funkce

## Statistické funkce

### Popisné statistiky

#### Aritmetický prùmìr

##### PRÙMÌR.W

Vypoèítá vážený aritmetický prùmìr. Oblast s hodnotami musí mít stejný rozsah (stejný poèet bunìk) jako oblast s váhami. Zadávejte bez záhlaví (pouze èísla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami, které mají být prùmìrovány |
| **váhy** | oblast bunìk s váhami |

 ```Excel
 =PRÙMÌR.W(A2:A10;B2:B10)
 ```

#### Harmonický prùmìr

##### HARMEAN.W

Vypoèítá vážený harmonický prùmìr. Oblast s hodnotami musí mít stejný rozsah (stejný poèet bunìk) jako oblast s váhami. Zadávejte bez záhlaví (pouze èísla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami, které mají být prùmìrovány |
| **váhy** | oblast bunìk s váhami |

 ```Excel
 =HARMEAN.W(A2:A10;B2:B10)
 ```

#### Geometrický prùmìr

##### GEOMEAN.W

Vypoèítá vážený geometrický prùmìr. Oblast s hodnotami musí mít stejný rozsah (stejný poèet bunìk) jako oblast s váhami. Zadávejte bez záhlaví (pouze èísla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bunìk s hodnotami, které mají být prùmìrovány |
| **váhy** | oblast bunìk s váhami |

 ```Excel
 =GEOMEAN.W(A2:A10;B2:B10)
 ```

#### Rozptyl a smìrodatná odchylka

##### VAR.P.W

Spoèítá vážený rozptyl pro populaci.

##### VAR.S.W

Spoèítá vážený rozptyl pro vzorek.

##### SMODCH.P.W

Spoèítá váženou smìrodatnou odchylku pro populaci.

##### SMODCH.S.W

Spoèítá váženou smìrodatnou odchylku pro vzorek.

##### VAR.RANGE

Spoèítá variaèní rozpìtí souboru.

##### VAR.RANGE

Spoèítá variaèní rozpìtí souboru.

### Distribuèní funkce

#### Normální rozdìlení

##### NORM.DIST.RANGE

Spoèítá pravdìpodobnost jevu mezi dvìma referenèními body u velièiny s normálním rozdìlením.

**argumenty**

| id | popis |
| --- | --- |
| **x** | strední hodnota rozdìlení |
| **s** | smìrodatná odchylka rozdìlení |
| **x1** | spodní mez |
| **x2** | horní mez |

### Analýza kontingenèní tabulky

#### KONTINGENCE.G

Vypoète testovou statistiku G z kontingenèní (køížové) tabulky. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.G(B2:D5)
 ```

#### KONTINGENCE.PV

Vypoète p-hodnotu pro kontingenèní (køížovou) tabulku. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.PV(B2:D5)
 ```

#### KONTINGENCE.C

Vypoète testovou statistiku Pearsonova C z kontingenèní (køížové) tabulky. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.C(B2:D5)
 ```

#### KONTINGENCE.V

Vypoète testovou statistiku Cramérovo V z kontingenèní (køížové) tabulky. Argumentem je výbìr tabulky bez záhlaví.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bunìk kontingenèní tabulky (bez záhlaví a souètù) |

 ```Excel
 =KONTINGENCE.V(B2:D5)
 ```

### Korelace

#### Spearmanùv korelaèní koeficient

##### SPEARMAN

Vypoèítá Spearmanùv korelaèní koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bunìk s hodnotami øady x |
| **hodnoty y** | oblast bunìk s hodnotami øady y |

##### SPEARMAN.T

Vypoèítá testovou statistiku T pro Spearmanùv korelaèní koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bunìk s hodnotami øady x |
| **hodnoty y** | oblast bunìk s hodnotami øady y |

##### SPEARMAN.PV

Vypoèítá p-hodnotu pro Spearmanùv korelaèní koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bunìk s hodnotami øady x |
| **hodnoty y** | oblast bunìk s hodnotami øady y |