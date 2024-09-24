# Funkce

Bal��ek roz���uje vzorce (funkce) Excelu o vybran� statistick� funkce, kter� urychluj� v�po�ty metod vyu�ovan�ch v p�edm�tech ST201 a ST210. Funkce obsahuj� z�kladn� dokumentaci v�stup� a vstupn�ch parametr�.

Doporu�uji za��t funkce za��t pou��vat bu�to tehdy, kdy� budete m�t dob�e zvl�dnut� ru�n� v�stupy, p��padn� jako kontrolu nad ru�n� vypo��tan�m v�sledkem.

**Popisn� statistiky**

| **n�zev funkce** | **popis funkce** |
| :--- | :--- |
| [MEAN.W](#V�en�-aritmetick�-pr�m�r) | v�en� aritmetick� pr�m�r |
| [HARMEAN.W](#V�en�-harmonick�-pr�m�r) | v�en� harmonick� pr�m�r |
| [GEOMEAN.W](#V�en�-geometrick�-pr�m�r) | v�en� geometrick� pr�m�r |
| [VAR.P.W](#V�en�-popula�n�-rozptyl)|	v�en� popula�n� rozptyl |
| [VAR.S.W](#V�en�-rozptyl-ze-vzorku) | v�en� rozptyl ze vzorku |
| [SMODCH.P.W](#V�en�-sm�rodatn�-odchylka-populace) | v�en� sm�rodatn� odchylka populace |
| [SMODCH.S.W](#V�en�-sm�rodatn�-odchylka-ze-vzorku) | v�en� sm�rodatn� odchylka ze vzorku |
| [VAR.RANGE](#Varia�n�-rozp�t�) | varia�n� rozp�t� souboru |
| [ROZKLAD.ROZPTYLU](#Rozklad-rozptylu) | rozklad rozptylu |
| [MAD](#Absolutn�-medi�nov�-odchylka) | absolutn� medi�nov� odchylka |

**Distribu�n� funkce**

| **n�zev funkce** | **popis funkce** |
| :--- | :--- |
| [NORM.DIST.RANGE](#Interval-norm�ln�ho-rozd�len�) | pravd�podobnost intervalu u norm�ln�ho rozd�len� |

**Anal�za kontingen�n� tabulky**

| **n�zev funkce** | **popis funkce** |
| :--- | :--- |
| [KONTINGENCE.G](#Testov�-statistika-G) | Testov� statistika G |
| [KONTINGENCE.PV](#P-hodnota-pro-kontingenci) | P-hodnota pro kontingenci |
| [KONTINGENCE.C](#Pearson�v-koeficient-kontingence-C) | Pearson�v koeficient kontingence C |
| [KONTINGENCE.V](#Cram�r�v-koeficient-kontingence-V) | Cram�r�v koeficient kontingence V |

**Spearman�v korela�n� koeficient**
| **n�zev funkce** | **popis funkce** |
| :--- | :--- |
| [SPEARMAN](#Spearman�v-korela�n�-koeficient) | Spearman�v korela�n� koeficient |
| [SPEARMAN.PV](#P-hodnota-pro-Spearman�v-test) | P-hodnota pro Spearmanovo rho |
| [SPEARMAN.T](#Testov�-statistika-T-pro-Sperman�v-test) | Pearson�v koeficient kontingence C |





## Popisn� statistiky

### V�en� aritmetick� pr�m�r

```Excel
MEAN.W
```

Vypo��t� v�en� aritmetick� pr�m�r. Oblast s hodnotami mus� m�t stejn� rozsah (stejn� po�et bun�k) jako oblast s v�hami. Zad�vejte bez z�hlav� (pouze ��sla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami, kter� maj� b�t pr�m�rov�ny |
| **v�hy** | oblast bun�k s v�hami |

 ```Excel
 =PR�M�R.W(A2:A10;B2:B10)
 ```

### V�en� harmonick� pr�m�r

```Excel
HARMEAN.W
```

Vypo��t� v�en� harmonick� pr�m�r. Oblast s hodnotami mus� m�t stejn� rozsah (stejn� po�et bun�k) jako oblast s v�hami. Zad�vejte bez z�hlav� (pouze ��sla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami, kter� maj� b�t pr�m�rov�ny |
| **v�hy** | oblast bun�k s v�hami |

 ```Excel
 =HARMEAN.W(A2:A10;B2:B10)
 ```

### V�en� geometrick� pr�m�r

```Excel
GEOMEAN.W
```

Vypo��t� v�en� geometrick� pr�m�r. Oblast s hodnotami mus� m�t stejn� rozsah (stejn� po�et bun�k) jako oblast s v�hami. Zad�vejte bez z�hlav� (pouze ��sla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami, kter� maj� b�t pr�m�rov�ny |
| **v�hy** | oblast bun�k s v�hami |

 ```Excel
 =GEOMEAN.W(A2:A10;B2:B10)
 ```

### V�en� popula�n� rozptyl

```Excel
VAR.P.W
```

Spo��t� v�en� rozptyl pro populaci.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s d�l��mi pr�m�ry |
| **v�hy** | oblast bun�k s v�hami |

```Excel
 =VAR.P.W(A2:A10;B2:B10)
```

### V�en� rozptyl ze vzorku

```Excel
VAR.S.W
```

Spo��t� v�en� rozptyl pro vzorek.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s d�l��mi pr�m�ry |
| **v�hy** | oblast bun�k s v�hami |

```Excel
=VAR.S.W(A2:A10;B2:B10)
```

### V�en� sm�rodatn� odchylka populace

```Excel
SMODCH.P.W
```

Spo��t� v�enou sm�rodatnou odchylku pro populaci.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s d�l��mi pr�m�ry |
| **v�hy** | oblast bun�k s v�hami |

```Excel
=SMODCH.P.W(A2:A10;B2:B10)
```

### V�en� sm�rodatn� odchylka ze vzorku

```Excel
SMODCH.S.W
```

Spo��t� v�enou sm�rodatnou odchylku pro vzorek.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s d�l��mi pr�m�ry |
| **v�hy** | oblast bun�k s v�hami |

```Excel
=SMODCH.S.W(A2:A10;B2:B10)
```

### Varia�n� rozp�t�

```Excel
VAR.RANGE
```

Spo��t� varia�n� rozp�t� souboru, tedy rozd�l mezi nejv�t�� a nejm�n�� hodnotou v souboru.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami |

```Excel
=VAR.RANGE(A2:A10)
```

### Rozklad rozptylu

```Excel
ROZKLAD.ROZPTYLU
```

Spo��t� v�tu o rozkladu rozptylu. Pro v�po�et je nutn� zadat d�l�� rozptyly, st�edn� hodnoty a v�hy.

**argumenty**

| id | popis |
| --- | --- |
| **rozptyly** | oblast bun�k s d�l��mi rozptyly |
| **pr�m�ry** | oblast bun�k s d�l��mi pr�m�ry |
| **v�hy** | oblast bun�k s v�hami |

```Excel
=ROZKLAD.ROZPTYLU(A2:A10;B2:B10;C2:C10)
```

### Absolutn� medi�nov� odchylka

```Excel
MAD
```

Spo��t� absolutn� medi�novou odchylku souboru.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami |

```Excel
=MAD(A2:A10)
```

## Distribu�n� funkce

### Interval norm�ln�ho rozd�len�

```Excel
NORM.DIST.RANGE
```

Spo��t� pravd�podobnost jevu mezi dv�ma referen�n�mi body u veli�iny s norm�ln�m rozd�len�m. Funguje podobn� jako funkce BINOM.DIST.RANGE.

**argumenty**

| id | popis |
| --- | --- |
| **x** | st�edn� hodnota rozd�len� |
| **s** | sm�rodatn� odchylka rozd�len� |
| **x1** | spodn� hranice |
| **x2** | horn� hranice |

```Excel
=NORM.DIST.RANGE(10;5;6;8)
```

## Anal�za kontingen�n� tabulky

### Testov� statistika G

```Excel
KONTINGENCE.G
```

Vypo�te testovou statistiku G z kontingen�n� (k��ov�) tabulky. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.G(B2:D5)
 ```

### P-hodnota pro kontingenci

```Excel
KONTINGENCE.PV
```

Vypo�te p-hodnotu pro kontingen�n� (k��ovou) tabulku. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.PV(B2:D5)
 ```

### Pearson�v koeficient kontingence C

```Excel
KONTINGENCE.C
```

Vypo�te testovou statistiku Pearsonova C z kontingen�n� (k��ov�) tabulky. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.C(B2:D5)
 ```

### Cram�r�v koeficient kontingence V

```Excel
KONTINGENCE.V
```

Vypo�te testovou statistiku Cram�rovo V z kontingen�n� (k��ov�) tabulky. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.V(B2:D5)
 ```

## Spearman�v korela�n� koeficient a testy

```Excel
SPEARMAN;
```

### Spearman�v korela�n� koeficient

Vypo��t� Spearman�v korela�n� koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bun�k s hodnotami �ady x |
| **hodnoty y** | oblast bun�k s hodnotami �ady y |

```Excel
=SPEARMAN(A2:A21;B2:B21)
```

### Testov� statistika T pro Sperman�v test

```Excel
SPEARMAN.T
```

Vypo��t� testovou statistiku T pro Spearman�v korela�n� koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bun�k s hodnotami �ady x |
| **hodnoty y** | oblast bun�k s hodnotami �ady y |

```Excel
=SPEARMAN.T(A2:A21;B2:B21)
```

### P-hodnota pro Spearman�v test

```Excel
SPEARMAN.PV
```

Vypo��t� p-hodnotu pro Spearman�v korela�n� koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bun�k s hodnotami �ady x |
| **hodnoty y** | oblast bun�k s hodnotami �ady y |

```Excel
=SPEARMAN.PV(A2:A21;B2:B21)
```