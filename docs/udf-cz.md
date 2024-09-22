# Funkce

## Statistick� funkce

### Popisn� statistiky

#### Aritmetick� pr�m�r

##### PR�M�R.W

Vypo��t� v�en� aritmetick� pr�m�r. Oblast s hodnotami mus� m�t stejn� rozsah (stejn� po�et bun�k) jako oblast s v�hami. Zad�vejte bez z�hlav� (pouze ��sla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami, kter� maj� b�t pr�m�rov�ny |
| **v�hy** | oblast bun�k s v�hami |

 ```Excel
 =PR�M�R.W(A2:A10;B2:B10)
 ```

#### Harmonick� pr�m�r

##### HARMEAN.W

Vypo��t� v�en� harmonick� pr�m�r. Oblast s hodnotami mus� m�t stejn� rozsah (stejn� po�et bun�k) jako oblast s v�hami. Zad�vejte bez z�hlav� (pouze ��sla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami, kter� maj� b�t pr�m�rov�ny |
| **v�hy** | oblast bun�k s v�hami |

 ```Excel
 =HARMEAN.W(A2:A10;B2:B10)
 ```

#### Geometrick� pr�m�r

##### GEOMEAN.W

Vypo��t� v�en� geometrick� pr�m�r. Oblast s hodnotami mus� m�t stejn� rozsah (stejn� po�et bun�k) jako oblast s v�hami. Zad�vejte bez z�hlav� (pouze ��sla).

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty** | oblast bun�k s hodnotami, kter� maj� b�t pr�m�rov�ny |
| **v�hy** | oblast bun�k s v�hami |

 ```Excel
 =GEOMEAN.W(A2:A10;B2:B10)
 ```

#### Rozptyl a sm�rodatn� odchylka

##### VAR.P.W

Spo��t� v�en� rozptyl pro populaci.

##### VAR.S.W

Spo��t� v�en� rozptyl pro vzorek.

##### SMODCH.P.W

Spo��t� v�enou sm�rodatnou odchylku pro populaci.

##### SMODCH.S.W

Spo��t� v�enou sm�rodatnou odchylku pro vzorek.

##### VAR.RANGE

Spo��t� varia�n� rozp�t� souboru.

##### VAR.RANGE

Spo��t� varia�n� rozp�t� souboru.

### Distribu�n� funkce

#### Norm�ln� rozd�len�

##### NORM.DIST.RANGE

Spo��t� pravd�podobnost jevu mezi dv�ma referen�n�mi body u veli�iny s norm�ln�m rozd�len�m.

**argumenty**

| id | popis |
| --- | --- |
| **x** | stredn� hodnota rozd�len� |
| **s** | sm�rodatn� odchylka rozd�len� |
| **x1** | spodn� mez |
| **x2** | horn� mez |

### Anal�za kontingen�n� tabulky

#### KONTINGENCE.G

Vypo�te testovou statistiku G z kontingen�n� (k��ov�) tabulky. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.G(B2:D5)
 ```

#### KONTINGENCE.PV

Vypo�te p-hodnotu pro kontingen�n� (k��ovou) tabulku. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.PV(B2:D5)
 ```

#### KONTINGENCE.C

Vypo�te testovou statistiku Pearsonova C z kontingen�n� (k��ov�) tabulky. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.C(B2:D5)
 ```

#### KONTINGENCE.V

Vypo�te testovou statistiku Cram�rovo V z kontingen�n� (k��ov�) tabulky. Argumentem je v�b�r tabulky bez z�hlav�.

**argumenty**

| id | popis |
| --- | --- |
| **observed** | oblast bun�k kontingen�n� tabulky (bez z�hlav� a sou�t�) |

 ```Excel
 =KONTINGENCE.V(B2:D5)
 ```

### Korelace

#### Spearman�v korela�n� koeficient

##### SPEARMAN

Vypo��t� Spearman�v korela�n� koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bun�k s hodnotami �ady x |
| **hodnoty y** | oblast bun�k s hodnotami �ady y |

##### SPEARMAN.T

Vypo��t� testovou statistiku T pro Spearman�v korela�n� koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bun�k s hodnotami �ady x |
| **hodnoty y** | oblast bun�k s hodnotami �ady y |

##### SPEARMAN.PV

Vypo��t� p-hodnotu pro Spearman�v korela�n� koeficient.

**argumenty**

| id | popis |
| --- | --- |
| **hodnoty x** | oblast bun�k s hodnotami �ady x |
| **hodnoty y** | oblast bun�k s hodnotami �ady y |