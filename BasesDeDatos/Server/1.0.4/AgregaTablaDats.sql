DROP TABLE IF EXISTS `dats`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `dats` (
  `id` int(11) NOT NULL DEFAULT '0',
  `nombre` varchar(255) DEFAULT '',
  `valor` text ,
  PRIMARY KEY (`id`),
  KEY `id` (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;



insert into dats (nombre,valor) values ('Hechizos','Tipos....
1 actuan sobre HP,MANA,STA,HAM y SED
2 actuan sobre los estados de los usuarios
4 invocacion
5 crea teleport
6 invoca familiar
7 materializa
8 Combinacion 1, 2 y 11
9 Actua sobre los estados de NPCs (Domar y Calmar)
10 Creación de arma,escudo, casco, armadura mágica
11 Modificación de estado de arma, armadura, escudo, casco
12 Detecta invisibles

Tagets....
1.....Usuario
2.....Npc
3.....Usuario Y Npc
4.....Terreno

[INIT]
NumeroHechizos=103

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO1]
Nombre=Curar veneno
Desc=Cura el envenenamiento
PalabrasMagicas=Nihil Ved
HechizeroMsg=Has curado a
TargetMsg=te ha curado del envenenamiento.
PropioMsg=Te has curado del envenenamiento.
Tipo=2
WAV=239
FXgrh=2
Loops=2
CuraVeneno=1
MinSkill=8
ManaRequerido=10
StaRequerido=1
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO2]
Nombre=Proyectil mágico
Desc=Causa 2 a 7 puntos daño a la víctima.
PalabrasMagicas=RahZaLu
HechizeroMsg=Has lanzado Proyectil mágico sobre
TargetMsg=lanzo Proyectil mágico sobre vos.
Tipo=1
WAV=233
Particle=102
SubeHP=2
MinHP=2
MaxHP=7
MinSkill=5
ManaRequerido=12
StaRequerido=2
Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO3]
Nombre=Curar heridas leves
Desc=Curar heridas leves, restaura entre 1 y 5 puntos de salud.
PalabrasMagicas=Corp Sanc

HechizeroMsg=Has sanado a
TargetMsg=te ha curado algunas heridas.
PropioMsg=Te has curado algunas heridas.

Tipo=1
WAV=238

SubeHP=1
MinHP=1
MaxHP=5

Particle=105

MinSkill=10
ManaRequerido=10
StaRequerido=3

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO4]
Nombre=Envenenar
Desc=Envenenamiento, provoca la muerte si no se contraresta con un antídoto.
PalabrasMagicas=SERP XON IN

HechizeroMsg=Has envenenado a
TargetMsg=te ha envenenado.
PropioMsg=Te has envenenado.

Tipo=2
WAV=16
Particle=32

Envenena=2
MinSkill=20
ManaRequerido=20
StaRequerido=4

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO5]
Nombre=Curar heridas graves
Desc=Curar heridas graves, restaura entre 15 y 25 puntos de salud.
PalabrasMagicas=EN CORP SANCTIS

HechizeroMsg=Has curado a
TargetMsg=te ha curado.
PropioMsg=Te has curado.

Tipo=1
WAV=237
FXgrh=9
loops=0

SubeHP=1
MinHP=15
MaxHP=25
    
MinSkill=35
ManaRequerido=40
StaRequerido=5

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO6]
Nombre=Saeta Ígnea
Desc=Causa 7 a 14 puntos de daño a la victima.
PalabrasMagicas=Fogûs Saex
HechizeroMsg=Has lanzado saeta ígnea sobre 
TargetMsg=lanzó Saeta ígnea sobre vos.
PropioMsg=Has lanzado Saeta ígnea sobre ti.
Tipo=1
WAV=19
Particle=96

SubeHP=2
MinHP=7
MaxHP=14
    
MinSkill=12
ManaRequerido=20
StaRequerido=5

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO7]
Nombre=Proyectil eléctrico
Desc=Causa 14 a 22 puntos de daño a la victima.
PalabrasMagicas=SUN VAP
HechizeroMsg=Has lanzado Proyectil eléctrico sobre
TargetMsg=lanzó Proyectil eléctrico sobre vos.
PropioMsg=Has lanzado proyectil eléctrico sobre ti.
Tipo=1
WAV=164
Particle=86

SubeHP=2
MinHP=14
MaxHP=22
    
MinSkill=22
ManaRequerido=30
StaRequerido=7

Target=3
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO8]
Nombre=Bomba mágica
Desc=Causa 10 a 30 puntos de daño en un área de dos tiles.
PalabrasMagicas=VAX IN TAR
HechizeroMsg=Has lanzado Bomba mágica sobre
TargetMsg=lanzó Bomba mágica sobre vos.
PropioMsg=Has lanzado Bomba mágica sobre ti.
Tipo=1
WAV=129
FXgrh=50
loops=1


SubeHP=2
MinHP=5
MaxHP=5

MinSkill=38
ManaRequerido=50
StaRequerido=20

Target=4
Afecta=3
HechizoDeArea=1
AreaEfecto=2

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO9]
Nombre=Paralizar
Desc=Impide que la víctima realize todo movimiento.
PalabrasMagicas=HOAX VORP

HechizeroMsg=Has paralizado a
TargetMsg=te ha paralizado.
PropioMsg=Te has paralizado.

Tipo=2
WAV=203
FXgrh=8
loops=0
Paraliza=1

MinSkill=60
ManaRequerido=350
StaRequerido=90

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO10]
Nombre=Remover paralisis
Desc=Remueve la paralisis.
PalabrasMagicas=AN HOAX VORP

HechizeroMsg=Le has removido la paralisis a
TargetMsg=te ha removido la paralisis.
PropioMsg=Te has removido la paralisis.

Tipo=2
WAV=255

RemoverParalisis=1

MinSkill=45
ManaRequerido=300
StaRequerido=80

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO11]
Nombre=Resucitar
Desc=Resucitar un usuario muerto.
PalabrasMagicas=AHIL KNÄ XÄR

HechizeroMsg=Has resucitado a
TargetMsg=te ha resucitado.
PropioMsg=Te has resucitado.

Tipo=2
WAV=240

Revivir=1

MinSkill=75
ManaRequerido=420
StaRequerido=100

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO12]
Nombre=Grito de Igôr
Desc=Aumenta criticamente la fuerza.
PalabrasMagicas=ÔL HÀX

HechizeroMsg=Le has lanzado Grito de Igôr a
TargetMsg=te ha lanzado el hechizo Grito de Igôr.
PropioMsg=Te has lanzado el hechizo Grito de Igôr.

Tipo=1
WAV=92


SubeFU=1
MinFU=15
MaxFU=15
    
MinSkill=55
ManaRequerido=360
StaRequerido=40

Target=1


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO13]
Nombre=Furia de Uhkrul
Desc=Furia de Uhkrul, aumenta la fuerza y la agilidad del objetivo criticamente.
PalabrasMagicas=ÛXÔLHIC

HechizeroMsg=Le has lanzado Furia de Uhkrul a
TargetMsg=te ha lanzado Furia de Uhkrul.
PropioMsg=Te has lanzado Furia de Uhkrul.

Tipo=1

WAV=102
Particle=111

SubeFU=1
MinFU=12
MaxFU=16

SubeAG=1
MinAG=12
MaxAG=16
    
MinSkill=85
ManaRequerido=720
StaRequerido=250

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO14]
Nombre=Invisibilidad
Desc=Vuelve invisible al objetivo, efecto no permanente.
HechizeroMsg=Has lanzado Invisibilidad sobre
TargetMsg=lanzo Invisibilidad sobre vos.
Tipo=2
WAV=248
FXgrh=46

Invisibilidad=1

MinSkill=87
ManaRequerido=560
StaRequerido=60

Target=1
Autolanzar=1


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO15]
Nombre=Tormenta de fuego
Desc=Causa 15 a 50 puntos de daño en un área de tres tiles.
PalabrasMagicas=RAX IN ZAR
HechizeroMsg=Has lanzado Tormenta de fuego sobre
TargetMsg=lanzo Tormenta de fuego sobre vos.
PropioMsg=Has lanzado Tormenta de fuego sobre ti.
Tipo=1
WAV=27
FXgrh=7
loops=0

SubeHP=2
MinHP=20
MaxHP=50
    
MinSkill=75
ManaRequerido=250
StaRequerido=60

HechizoDeArea=1
AreaEfecto=3
Afecta=3

Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO16]
Nombre=Llamado musical
Desc=Implora el espiritu musical, un bardo acudirá en tu ayuda.
PalabrasMagicas=Impendere et worg
HechizeroMsg=Has invocado un bardo.

Tipo=4
WAV=38
Fxgrh=17

Invoca=1
NumNpc=163
Cant=1

MinSkill=50
ManaRequerido=460
StaRequerido=35

Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO17]
Nombre=Llamado a la naturaleza
Desc=Trae tres lobos que ayudarán a tu causa.
PalabrasMagicas=Lüpus Aident
HechizeroMsg=Has invocado tres lobos.

Tipo=4
WAV=124
Particle=78

Invoca=1
NumNpc=545
Cant=3
MinSkill=30
ManaRequerido=250
StaRequerido=40
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO18]
Nombre=Celeridad
Desc=Aumenta la agilidad del usuario que recibe el spell
PalabrasMagicas=YUP AINC
HechizeroMsg=Has lanzado Celeridad sobre 
TargetMsg=ha lanzado Celeridad sobre ti.
PropioMsg=Has lanzado Celeridad sobre ti.

Tipo=1
WAV=230

SubeAG=1
MinAG=2
MaxAG=5

MinSkill=23
ManaRequerido=60
StaRequerido=20
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO19]
Nombre=Torpeza
Desc=Reduce la agilidad del usuario que recibe el spell
PalabrasMagicas=ASYNC YUP AINC
HechizeroMsg=Has lanzado Torpeza sobre 
TargetMsg=ha lanzado Torpeza sobre ti.
PropioMsg=Has lanzado Torpeza sobre ti.

Tipo=1
WAV=17

SubeAG=2
MinAG=2
MaxAG=5

MinSkill=30
ManaRequerido=40
StaRequerido=10
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO20]
Nombre=Fuerza
Desc=Aumenta la fuerza del objetivo
PalabrasMagicas=Ar Akron
HechizeroMsg=Has lanzado Fuerza sobre 
TargetMsg=ha lanzado Fuerza sobre ti.
PropioMsg=Has lanzado Fuerza sobre ti.

Tipo=1
WAV=230

SubeFU=1
MinFU=2
MaxFU=5

MinSkill=35
ManaRequerido=75
StaRequerido=25
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO21]
Nombre=Debilidad
Desc=Reduce la fuerza del usuario que recibe el spell
PalabrasMagicas=Xoom Varp
HechizeroMsg=Has lanzado Debilidad sobre 
TargetMsg=ha lanzado Debilidad sobre ti.
PropioMsg=Has lanzado Debilidad sobre ti.

Tipo=1
WAV=17

SubeFU=2
MinFU=2
MaxFU=5

MinSkill=35
ManaRequerido=50
StaRequerido=15
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO22]
Nombre=Desencantar
Desc=Desencantar remueve todo efecto mágico del usuario que lo conjure.
PalabrasMagicas=Disincanto
HechizeroMsg=Has lanzado Desencantar sobre
TargetMsg=lanzo Desencantar sobre vos.
PropioMsg=Has lanzado Desencantar sobre ti.
Tipo=2
WAV=59
Particle=85
desencantar=1
autolanzar=1

MinSkill=100
ManaRequerido=1250
StaRequerido=250
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO23]
Nombre=Corriente eléctrica
Desc=Causa 20 a 70 puntos de daño en un área de cuatro tiles.
PalabrasMagicas=THY KOOOL
HechizeroMsg=Has lanzado Corriente eléctrica sobre
TargetMsg=lanzo Corriente eléctrica sobre vos.
PropioMsg=Has lanzado Corriente eléctrica sobre ti.
Tipo=1
WAV=56
Particle=114
loops=1

SubeHP=2
MinHP=55
MaxHP=70
    
MinSkill=75
ManaRequerido=300
StaRequerido=70

HechizoDeArea=1
AreaEfecto=4
Afecta=3

Target=2


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO24]
Nombre=Inmovilizar
Desc=Impide que la víctima se movilize.
PalabrasMagicas=Är Prop suo

HechizeroMsg=Has inmovilizado a
TargetMsg=te ha inmovilizado.
PropioMsg=Te has inmovilizado.

Tipo=2
WAV=253
FXgrh=12
loops=0
Inmoviliza=1

MinSkill=55
ManaRequerido=230
StaRequerido=60
Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO25]
Nombre=Apocalipsis
Desc=Causa 102 a 124 puntos de daño a la victima.
PalabrasMagicas=Rahma Nañarak Oal
HechizeroMsg=Has lanzado Apocalipsis sobre
TargetMsg=lanzo Apocalipsis sobre vos.
PropioMsg=Has lanzado Apocalipsis sobre ti.
ExclusivoClase=MAGO-NIGROMANTE
Tipo=1
WAV=138
Particle=45

SubeHP=2
MinHP=102
MaxHP=124
    
MinSkill=80
ManaRequerido=1050
StaRequerido=110

Target=3

Anillo=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO26]
Nombre=Invocar elemental de fuego
Desc=Invocar elemental de fuego
PalabrasMagicas=Yurrax
HechizeroMsg=Has invocado un elemental de fuego.
Tipo=4
WAV=28
FXgrh=7
Invoca=1
NumNpc=93
Cant=1
MinSkill=100
ManaRequerido=740
StaRequerido=100
ExclusivoClase=MAGO-NIGROMANTE-CLERIGO-DRUIDA-BARDO
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO27]
Nombre=Invocar elemental de agua
Desc=Invocar elemental de ague
PalabrasMagicas=Mantrarax
HechizeroMsg=Has invocado un elemental de agua.
Tipo=4
WAV=28
Particle=116
Invoca=1
NumNpc=92
Cant=1
MinSkill=90
ManaRequerido=780
StaRequerido=110
ExclusivoClase=MAGO-NIGROMANTE-CLERIGO-DRUIDA-BARDO
Target=4


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO28]
Nombre=Invocar elemental de tierra
Desc=Invocar elemental de tierra
PalabrasMagicas=Mantrarax
HechizeroMsg=Has invocado un elemental de tierra.
Tipo=4
WAV=28
Invoca=1
NumNpc=94
Cant=1
MinSkill=100
ManaRequerido=800
StaRequerido=100
ExclusivoClase=MAGO-NIGROMANTE-CLERIGO-DRUIDA-BARDO
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO29]
Nombre=Implorar ayuda
Desc=Implorar ayuda
PalabrasMagicas=ArCos Mantrarax
HechizeroMsg=¡Has implorado ayuda a los dioses!
Tipo=4
WAV=17
FXgrh=7
Invoca=1
NumNpc=146
Cant=1
MinSkill=100
ManaRequerido=2000
StaRequerido=240
ExclusivoClase=MAGO-NIGROMANTE-CLERIGO-DRUIDA-BARDO
Target=4


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO33]
Nombre=Desatar la ira
Desc=Desatar la ira reduce la fuerza y la agilidad del usuario críticamente.
PalabrasMagicas=Hic Dei

HechizeroMsg=Has desatado la ira de
TargetMsg=te ha desatado la ira
PropioMsg=Te has desatado la ira.

Tipo=1
WAV=32
Particle=112

SubeFU=2
MinFU=5
MaxFU=10

SubeAG=2
MinAG=5
MaxAG=10

MinSkill=90
ManaRequerido=1440
StaRequerido=250

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
[HECHIZO34]
Nombre=Implosión
Desc=Causa 125 a 138 puntos de daño a la victima.
PalabrasMagicas=Nihil Mortem
HechizeroMsg=Has lanzado Implosión sobre
TargetMsg=lanzo Implosión sobre vos.
PropioMsg=Has lanzado Implosión sobre ti.
Tipo=1
WAV=65
Fxgrh=18
ExclusivoClase=MAGO-NIGROMANTE-CLERIGO-DRUIDA-BARDO
loops=0

SubeHP=2
MinHP=125
MaxHP=138
    
MinSkill=100
ManaRequerido=1500
StaRequerido=120

Anillo=1

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO35]
Nombre=Lluvia de Mosquitos
Desc=Los insectos paralizarán y dejaran ciega a la victima.
PalabrasMagicas=Hut Inefectus
HechizeroMsg=Has lanzado lluvia de mosquitos sobre
TargetMsg=lanzó Lluvia de mosquitos sobre vos.
PropioMsg=Has lanzado Lluvia de mosquitos sobre ti.
Tipo=2
WAV=26
Particle=115
loops=9

Paraliza=1
Ceguera=1

MinSkill=100
ManaRequerido=2500
StaRequerido=300
ExclusivoClase=MAGO

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO36]  Elemental de fuego, meduzas, etc
Nombre=Hechizo de NPC
Tipo=1
WAV=27
FXgrh=7
loops=0

SubeHP=2
MinHP=30
MaxHP=50
Target=3

[HECHIZO37]
Nombre=HechizoNPCB
Desc=Hechizo de NPC.
Tipo=1
WAV=10
FXgrh=14
loops=0

SubeHP=2
MinHP=90
MaxHP=120
Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO38]
Nombre=Metamorfosis: Gólem
Desc=Transforma al usuario en un Gólem, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en gólem a
TargetMsg=te transformó en gólem.
PropioMsg=Te has transformado en gólem.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=371
head=0
MinSkill=100
ManaRequerido=1000
StaRequerido=200
Target=1
MetaObj=1605
ExtraHIT=30
ExtraDEF=30
ExclusivoClase=DRUIDA
autolanzar=1

[HECHIZO39]
Nombre=Metamorfosis: Dragón
Desc=Transforma al usuario en un Dragón, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en dragón a
TargetMsg=te transformó en dragón.
PropioMsg=Te has transformado en dragón.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=357
head=0
ExtraHIT=50
ExtraDEF=20
MinSkill=100
ManaRequerido=1000
StaRequerido=200
Target=1
MetaObj=1604
vuela=1
autolanzar=1

[HECHIZO40]
Nombre=Metamorfosis: Lobo
Desc=Transforma al usuario en un Lobo, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en lobo a
TargetMsg=te transformó en lobo.
PropioMsg=Te has transformado en lobo.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=10
head=0
ExtraHIT=10
ExtraDEF=15
MinSkill=15
ManaRequerido=100
StaRequerido=20
Target=1
MetaObj=1611
ExclusivoClase=DRUIDA
autolanzar=1

[HECHIZO41]
Nombre=Metamorfosis: Ogro
Desc=Transforma al usuario en un Ogro, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en ogro a
TargetMsg=te transformó en ogro.
PropioMsg=Te has transformado en ogro.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=76
head=0

ExtraHIT=30
ExtraDEF=10
MinSkill=70
ManaRequerido=600
StaRequerido=40
Target=1
MetaObj=0
autolanzar=1

[HECHIZO42]
Nombre=Metamorfosis: Jabalí
Desc=Transforma al usuario en un Jabalí, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en jabalí a
TargetMsg=te transformó en jabalí.
PropioMsg=Te has transformado en jabalí.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=406
head=0
ExtraHIT=10
ExtraDEF=15
MinSkill=15
ManaRequerido=100
StaRequerido=30
Target=1
MetaObj=0
autolanzar=1

[HECHIZO43]
Nombre=Metamorfosis: Árbol
Desc=Transforma al usuario en un Árbol, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en árbol a
TargetMsg=te transformó en árbol.
PropioMsg=Te has transformado en árbol.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=24
head=0
ExtraHIT=6
ExtraDEF=6
MinSkill=90
ManaRequerido=900
StaRequerido=100
Target=1
MetaObj=0
autolanzar=1

[HECHIZO44]
Nombre=Metamorfosis: Oso Pardo
Desc=Transforma al usuario en un Oso Pardo, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en oso pardo a
TargetMsg=te transformó en oso pardo.
PropioMsg=Te has transformado en oso pardo.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=73
head=0
ExtraHIT=20
ExtraDEF=8
MinSkill=20
ManaRequerido=200
StaRequerido=40
Target=1
MetaObj=1607
ExclusivoClase=DRUIDA
autolanzar=1

[HECHIZO45]
Nombre=Metamorfosis: Tortuga
Desc=Transforma al usuario en una Tortuga, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en tortuga a
TargetMsg=te transformó en tortuga.
PropioMsg=Te has transformado en tortuga.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=74
head=0
ExtraHIT=30
ExtraDEF=6
MinSkill=30
ManaRequerido=200
StaRequerido=40
Target=1
MetaObj=1610
ExclusivoClase=DRUIDA
autolanzar=1

[HECHIZO46]
Nombre=Metamorfosis: Araña
Desc=Transforma al usuario en una Araña, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en araña a
TargetMsg=te transformó en araña.
PropioMsg=Te has transformado en araña.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=42
head=0
ExtraHIT=25
ExtraDEF=10
MinSkill=50
ManaRequerido=300
StaRequerido=60
Target=1
MetaObj=0
autolanzar=1

[HECHIZO47]
Nombre=Metamorfosis: Demonio
Desc=Transforma al usuario en un Demonio, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en demonio a
TargetMsg=te transformó en demonio.
PropioMsg=Te has transformado en demonio.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=365
head=0
ExtraHIT=25
ExtraDEF=15
MinSkill=70
ManaRequerido=600
StaRequerido=200
Target=1
MetaObj=0
autolanzar=1

[HECHIZO48]
Nombre=Metamorfosis: Ent
Desc=Transforma al usuario en una Ent, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en ent a
TargetMsg=te transformó en ent.
PropioMsg=Te has transformado en ent.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=273
head=0
ExtraHIT=8
ExtraDEF=16
MinSkill=5
ManaRequerido=1000
StaRequerido=100
Target=1
MetaObj=1608
autolanzar=1

[HECHIZO49]
Nombre=Metamorfosis: Ágila
Desc=Transforma al usuario en un Ágila, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en ágilao a
TargetMsg=te transformó en ágila.
PropioMsg=Te has transformado en ágila.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=59
head=0
ExtraHIT=6
ExtraDEF=20
MinSkill=5
ManaRequerido=100
StaRequerido=30
Target=1
MetaObj=1606
autolanzar=1

[HECHIZO50]
Nombre=Metamorfosis: Tigre
Desc=Transforma al usuario en un Tigre, efecto no permanente. Las criaturas no te atacarán si aparentas ser más poderoso que ellas, y obtendrás un bonificador en tu ataque y defensa.
HechizeroMsg=Has transformado en tigre a
TargetMsg=te transformó en tigre.
PropioMsg=Te has transformado en tigre.
Tipo=2
WAV=57
fxgrh=30
metamorfosis=1
body=147
head=0
ExtraHIT=10
ExtraDEF=15
MinSkill=15
ManaRequerido=100
StaRequerido=30
Target=1
MetaObj=1609
ExclusivoClase=DRUIDA
autolanzar=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO51]
Nombre=HechizoNPCC
Desc=Hechizo de NPC.
Tipo=1
WAV=56
FXgrh=11
loops=1

SubeHP=2
MinHP=60
MaxHP=100

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO52]
Nombre=Juicio final
Desc=Causa 135 a 150 puntos de daño a la victima.
PalabrasMagicas=Finis Mortem
HechizeroMsg=Has lanzado Juicio final sobre
TargetMsg=lanzo Juicio final sobre vos.
PropioMsg=Has lanzado Juicio final sobre ti.
Tipo=1
WAV=66
FXgrh=20
loops=0


SubeHP=2
MinHP=135
MaxHP=150
    
MinSkill=100
ManaRequerido=1800
StaRequerido=200

Anillo=1
ExclusivoClase=MAGO-NIGROMANTE
Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO53]
Nombre=Portal planar
Desc=El hechizo Portal planar crea un teleport hacia Intermundia.
PalabrasMagicas=Opercure Dimentia
HechizeroMsg=Has creado un Portal interplanar.
Tipo=5
WAV=249
Particle=97
loops=0
    
MinSkill=100
ManaRequerido=2000
StaRequerido=300
ExclusivoClase=MAGO-NIGROMANTE
Target=4


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO55]
Nombre=Materializar: Comida
Desc=Los hechizos de materialización convierten magia en materia. Este hechizo puede crear comida a partir de su lanzamiento.
PalabrasMagicas=Sunt Alifens
HechizeroMsg=Has creado algo de comida.

Tipo=7
WAV=130
Particle=94
loops=1

MinSkill=90
ManaRequerido=500
StaRequerido=120
IndiceDeItem=1
ExclusivoClase=MAGO

Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO56]
Nombre=Incinerar
Desc=Una llamarada incinerará al objetivo, causando un daño considerable rápidamente.
PalabrasMagicas=Incinera Corps

HechizeroMsg=Has incinerado a
TargetMsg=te ha incinerado.
PropioMsg=Te has incinerado.

Tipo=2
WAV=123
Particle=96


Incinera=1
MinSkill=90
ManaRequerido=1500
StaRequerido=100
ExclusivoClase=MAGO

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO57]
Nombre=Tormenta ácida
Desc=Causa 20 a 40 puntos de daño y envenena en un área de tres tiles.
PalabrasMagicas=VAX IN XORP
HechizeroMsg=Has lanzado Tormenta ácida sobre
TargetMsg=lanzo Tormenta ácida sobre vos.
PropioMsg=Has lanzado Tormenta ácida sobre ti.
Tipo=1
WAV=127
FXgrh=32
loops=0

SubeHP=2
MinHP=20
MaxHP=40
    
MinSkill=80
ManaRequerido=400
StaRequerido=100
Envenena=3
ExclusivoClase=MAGO-NIGROMANTE-CLERIGO-DRUIDA-BARDO-PALADIN

HechizoDeArea=1
AreaEfecto=3
Afecta=3

Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO58]
Nombre=Maldición
Desc=Impide que la víctima ataque.
PalabrasMagicas=Malefectus Sunt
HechizeroMsg=Has maldecido a
TargetMsg=te ha maldecido.
PropioMsg=Te has maldecido.
Tipo=1
WAV=16
FXgrh=24
loops=0

SubeHP=2
MinHP=33
MaxHP=45

NoPuedeAtacar=1

MinSkill=55
ManaRequerido=1800
StaRequerido=100

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO59]
Nombre=Llamado al familiar
Desc=Llama o retira tu familiar
PalabrasMagicas=Familiarê Asimis
Tipo=6
WAV=246
FXgrh=7
Invoca=1
Cant=1
MinSkill=0
ManaRequerido=0
StaRequerido=0
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO60]
Nombre=Materializar: Bebida
Desc=Los hechizos de materialización convierten magia en materia. Este hechizo puede crear bebida a partir de su lanzamiento.
PalabrasMagicas=Sunt Aquum
HechizeroMsg=Has creado algo de bebida.

Tipo=7
WAV=130
Particle=94
loops=1

MinSkill=90
ManaRequerido=500
StaRequerido=120
IndiceDeItem=43
ExclusivoClase=MAGO
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO62]
Nombre=Detectar invisibilidad
Desc=Detecta todo usuario invisible en el rango de visión del usuario.
HechizeroMsg=Has lanzado Detectar invisibilidad.
PalabrasMagicas=Fant Visiblî

Target=4
Tipo=12
WAV=254
Particle=82
loops=1

MinSkill=95
ManaRequerido=1800
ExclusivoClase=MAGO
StaRequerido=300

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO63]
Nombre=Lamento de la Banshee
Desc=Causa 120 a 145 puntos de daño en un área de ocho tiles.
PalabrasMagicas=VÂR NI MÈS
HechizeroMsg=Has lanzado Lamento de la Banshee sobre
TargetMsg=lanzo Lamento de la Banshee sobre vos.
PropioMsg=Has lanzado Lamento de la Banshee sobre ti.
Tipo=1
WAV=226
FXgrh=45
loops=0

SubeHP=2
MinHP=120
MaxHP=145
    
MinSkill=100
ManaRequerido=1080
StaRequerido=180
ExclusivoClase=MAGO

Anillo=2

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO64]
Nombre=Veneno crítico
Desc=Envenenamiento crítico, provoca la muerte rápidamente si no se contraresta con un antídoto.
PalabrasMagicas=Pestis Xôn In

HechizeroMsg=Has envenenado críticamente a
TargetMsg=te ha envenenado críticamente.
PropioMsg=Te has envenenado críticamente.

Tipo=2
WAV=127

Envenena=5
MinSkill=60
ManaRequerido=250
StaRequerido=90

Target=3


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&CLERIGOS&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO32]
Nombre=Sanar
Desc=Restaura entre 60 y 100 de daño al usuario y remueve todos los efectos negativos que éste pueda tener.
PalabrasMagicas=Nihil Vitae
HechizeroMsg=Has lanzado Sanar sobre
TargetMsg=lanzó Sanar sobre vos.
PropioMsg=Has lanzado Sanar sobre ti.
Tipo=1
WAV=236
Particle=119
loops=0
Sanacion=1

SubeHP=1
MinHP=60
MaxHP=100
    
MinSkill=100
ManaRequerido=1000
StaRequerido=550
ExclusivoClase=Clerigo

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO54]
Nombre=Resurrección
Desc=Revive completamente a un usuario.
PalabrasMagicas=HÎC KNÄ XÄR

HechizeroMsg=Has revivido a
TargetMsg=te ha revivido.
PropioMsg=Te has revivido.

Tipo=2
WAV=256
FXgrh=0

Resurreccion=1

MinSkill=100
ManaRequerido=840
StaRequerido=200
ExclusivoClase=CLERIGO

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO67]
Nombre=Curar heridas críticas
Desc=Curar heridas críticas, restaura entre 30 y 60 puntos de salud.
PalabrasMagicas=CORP CURATIO

HechizeroMsg=Has curado a
TargetMsg=te ha curado.
PropioMsg=Te has curado.

Tipo=1
WAV=18
Particle=118

SubeHP=1
MinHP=30
MaxHP=60
    
MinSkill=65
ManaRequerido=120
StaRequerido=80
ExclusivoClase=CLERIGO

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO68] Lvl 25 en adelante 
Nombre=Palabra sagrada
Desc=La palabra sagrada paraliza y debilita al objetivo.
PalabrasMagicas=Feresdeth Averin

HechizeroMsg=Has lanzado Palabra sagrada sobre
TargetMsg=te ha lanzado Palabra sagrada
PropioMsg=Te has lanzado Palabra sagrada

Tipo=2
Target=1
WAV=
FXgrh=

Paraliza=1

SubeFU=2
MinFU=25
MaxFU=30

MinSkill=87
ManaRequerido=900
StaRequerido=180
ExclusivoClase=CLERIGO

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO69] Lvl 32 en adelante
Nombre=Descarga flamígera
Desc=Castiga al enemigo con el fuego divino, causando entre 70 y 110 puntos de daño e incinerándolo
PalabrasMagicas=Jâf Rashida

HechizeroMsg=Has lanzado Descarga Flamígera sobre
TargetMsg=te ha lanzado Descarga Flamígera
PropioMsg=Te has lanzado Descarga Flamígera

Tipo=2
WAV=242
Particle=84
loops=0

SubeHP=2
MinHP=70
MaxHP=110
    
MinSkill=100
ManaRequerido=950
StaRequerido=150
ExclusivoClase=CLERIGO

Anillo=1

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO70] Lvl 20 en adelante
Nombre=Aura sagrada
Desc=Los dioses te resguardan por medio de un aura protectora
PalabrasMagicas=Elinethik
PropioMsg=Un aura te protege

Tipo=10
Target=1
Wav=104
Particle=27

Creatipo=2

Anim=1
MinDef=4
MaxDef=12

MinSkill=50
ManaRequerido=1000
StaRequerido=300
ExclusivoClase=CLERIGO

autolanzar=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO71] Lvl 28 en adelante
Nombre=Pregonar la palabra de Dios
Desc=Cura en un diámetro de 5 tiles entre 75 y 125 puntos de vida
PalabrasMagicas=In nómini Pater, Filum, Sanctis Espiritum

HechizeroMsg=La palabra de Dios cura y fortalece
TargetMsg=te ha curado.
PropioMsg=Te has curado.

Tipo=1
WAV=99
Particle=101
loops=1

SubeHP=1
MinHP=75
MaxHP=125
    
MinSkill=70
ManaRequerido=1000
StaRequerido=350

AreaEfecto=5
HechizoDeArea=1

Target=4 Efecto área
ExclusivoClase=CLERIGO
Afecta=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO72] Lvl 42 en adelante
Nombre=Golpe certero
Desc=El receptor no fallará el próximo golpe
PalabrasMagicas=ArSen Benerim
HechizeroMsg=Has lanzado Golpe certero sobre
Targetmsg=te ha lanzado Golpe certero
PropioMsg=¡Tu próximo golpe no fallará!

Tipo=2
WAV=245

GolpeCertero=1
Particle=70

MinSkill=90
ManaRequerido=1480
StaRequerido=300
ExclusivoClase=CLERIGO

Target=1

autolanzar=1
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO73] Lvl 35 en adelante
Nombre=Poder divino
Desc=El poder del receptor aumenta considerablemente
PalabrasMagicas=Anath Ekelnar
HechizeroMsg=Has lanzado Poder Divino sobre
Targetmsg=te ha lanzado Poder Divino
PropioMsg=Te sientes más fuerte

Tipo=8
WAV=228

Particle=103

SubeFU=1
MinFU=15
MaxFU=25

SubeAG=1
MinAG=15
MaxAG=25

MinSkill=90
ManaRequerido=1260
StaRequerido=250
ExclusivoClase=CLERIGO

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&DRUIDAS&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


[HECHIZO74] Lvl 28 en adelante
Nombre=Animar a la naturaleza
Desc=Le da vida momentáneamente a un árbol.
PalabrasMagicas=In amari söpra Tai
HechizeroMsg=La naturaleza te protege

Tipo=4
WAV=17
Particle=100
Invoca=1
NumNpc=155
Cant=1
MinSkill=70
ManaRequerido=1000
StaRequerido=220
ExclusivoClase=DRUIDA
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO75] Lvl 9 en adelante
Nombre=Círculo curativo
Desc=Curar entre 25 y 35 puntos de salud en un diámetro de 3 tiles.
PalabrasMagicas=AmaÎri Sanctis

HechizeroMsg=Has lanzado círculo curativo sobre
TargetMsg=te ha curado.
PropioMsg=Te has curado.

Tipo=1
WAV=18

SubeHP=1
MinHP=25
MaxHP=35

Particle=106
    
MinSkill=45
ManaRequerido=300
StaRequerido=90

hechizodearea=1
areaefecto=3
ExclusivoClase=DRUIDA

Target=4 Efecto área
Afecta=1


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO76] Lvl 30 en adelante
Nombre=Semillas de fuego
Desc=Bellotas y bayas se convierten en bolas de fuego, causando 110 a 125 puntos de daño.
PalabrasMagicas=Samarin Tar
HechizeroMsg=Has lanzado Semillas de fuego sobre
TargetMsg=lanzo Semillas de fuego sobre vos.
PropioMsg=Has lanzado Semillas de fuego sobre ti.
Tipo=1
WAV=27
Particle=204
loops=0

SubeHP=2
MinHP=110
MaxHP=125
    
MinSkill=100
ManaRequerido=1250
StaRequerido=100
ExclusivoClase=DRUIDA

Anillo=1

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO77] Lvl 35/36 en adelante
Nombre=Rayo solar
Desc=Haz de luz que ciega e inflige 70 a 100 puntos de daño.
PalabrasMagicas=Pelor In bodem
HechizeroMsg=Has lanzado Rayo solar sobre
TargetMsg=lanzo Rayo solar sobre vos.
PropioMsg=Has lanzado Rayo solar sobre ti.
Tipo=1
WAV=27
Particle=98
loops=0

SubeHP=2
MinHP=70
MaxHP=100

Ceguera=1
    
MinSkill=100
ManaRequerido=1350
StaRequerido=125
ExclusivoClase=DRUIDA

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO78] Lvl 33 en adelante
Nombre=Terremoto
Desc=Causa entre 50 y 75 puntos de daño en un radio de 3 tiles.
PalabrasMagicas=Nature Obad-Hai Eventî

HechizeroMsg=Has lanzado un terremoto
TargetMsg=te ha atrapado en el terremoto.
PropioMsg=Te has atrapado en el terremoto.

Tipo=1
WAV=247

SubeHP=2
MinHP=50
MaxHP=75
    
MinSkill=95
ManaRequerido=650
StaRequerido=200

Particle=44

Afecta=3
hechizodearea=1
areaefecto=3
ExclusivoClase=DRUIDA
Target=4 Efecto área

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO79]
Nombre=Resucitar familiar o mascota
Desc=Resucita el familiar o mascota del objetivo.
PalabrasMagicas=Nature Xâr

TargetMsg=ha resucitado tu familiar.
PropioMsg=Has resucitado tu familiar.

Tipo=2
WAV=55
ResucitaFamiliar=1

MinSkill=70
ManaRequerido=1200
StaRequerido=120
ExclusivoClase=DRUIDA

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&BARDOS&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO82] lvl 32 en adelante
Nombre=Intimidar al espíritu
Desc=Reduce la fuerza y el daño del arma del receptor
PalabrasMagicas=Espiritus Exterrere
HechizeroMsg=Has intimidado a
TargetMsg=te ha intimidado
PropioMsg=Te has intimidado

Tipo=1
WAV=17

Particle=111

SubeFU=2
MinFU=25
MaxFU=35

CreaTipo=3
MinDef=-20
MaxDef=-20

StaRequerido=200
MinSkill=80
ManaRaquerido=900

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO83] Lvl 30 en adelante
Nombre=Cobijo seguro de Leomundo
Desc=Una fuerza invisible te otorga defensa
PalabrasMagicas=Leomund secure cobijim
HechizeroMsg=Te sientes protegido

Tipo=1
WAV=92

Particle=109

SubeFU=1
MinFU=15
MaxFU=15

CreaTipo=3
MinDef=+40
MaxDef=+40
    
MinSkill=90
ManaRequerido=1000
StaRequerido=300
ExclusivoClase=BARDO

autolanzar=1

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO84] Lvl 45 en adelante [EDIT deberia ser catidad / 100, en ese lvl debe tener como 2000 muertes]
Nombre=Lamento de las almas
Desc=Causa entre 110 y 125 puntos de daño
PalabrasMagicas=Espirictum Nifis Mortem
HechizeroMsg=Has lanzado Lamento de las almas sobre
TargetMsg=lanzo Lamento de las almas sobre vos.
PropioMsg=Has lanzado Lamento de las almas sobre ti.
Tipo=1
WAV=98
Particle=112

SubeHp=2
MinHp=110
MaxHp=125
ExclusivoClase=BARDO

MinSkill=100
ManaRequerido=950
StaRequerido=115

Anillo=1
Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO85] lvl 25 en adelante
Nombre=Inspiracion
Desc=Cura malos estados y aumenta la agilidad
PalabrasMagicas=Inspiratum
HechizeroMsg=Has inspirado a
TargetMsg=te ha inspirado
PropioMsg=Te has inspirado

Tipo=1

WAV=30
Particle=20
loops=0

SubeFU=1
MinFU=12
MaxFU=16

SubeAG=1
MinAG=12
MaxAG=16

Sanacion=1
    
MinSkill=100
ManaRequerido=1000
StaRequerido=150
ExclusivoClase=BARDO

Target=1



&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&PALADINES&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


[HECHIZO87]
Nombre=Castigo Divino
Desc=Causa 75 a 100 puntos de daño a la victima.
PalabrasMagicas=Abra Ainum
HechizeroMsg=Has lanzado Castigo Divino sobre
TargetMsg=lanzo Castigo Divino vos.
PropioMsg=Has lanzado Castigo Divino sobre ti.
Tipo=1
WAV=240
Particle=86
Resis=1

SubeHP=2
MinHP=75
MaxHP=100
    
MinSkill=80
ManaRequerido=520
StaRequerido=100
ExclusivoClase=PALADIN

target=3


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO88] Nivel 25 en adelante
Nombre=Fuente de vida
Desc=Cura entre 30 y 45 puntos de daño al objetivo
PalabrasMagicas=Abra Ainum
HechizeroMsg=Has lanzado Fuente de Vida sobre
TargetMsg=lanzo Fuente de vida sobre vos.
PropioMsg=Has lanzado Fuente de Vida sobre ti.
Tipo=1
WAV=100
Particle=100
loops=0

SubeHP=1
MinHP=30
MaxHP=45
    
MinSkill=75
ManaRequerido=300
StaRequerido=75
ExclusivoClase=PALADIN

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO89] Lvl 42 en adelante. Hechizo powa
Nombre=Honor del Paladín
Desc=Los dioses te bendicen aumentando tu potencial
PalabrasMagicas=ArCos Sanctis Hic
PropioMsg=Te sientes más fuerte

Tipo=8
WAV=243
Particle=107
loops=0

SubeFU=1
MinFU=35
MaxFU=35

SubeAG=1
MinAG=35
MaxAG=35

MinSkill=100
ManaRequerido=700
StaRequerido=300
ExclusivoClase=PALADIN

AutoLanzar=1

Target=1

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO91] Lvl 35 en adelante
Nombre=Arma mágica
Desc=Crea momentáneamente un arma que inflige un daño considerable
PalabrasMagicas=Sarapiens Herinoenous
PropioMsg=Un arma mágica se blande en tu mano

Tipo=10
Target=1

Creatipo=1
Anim=1
Particle=53
loops=0
Wav=57

MinHit=14
MaxHit=21

AutoLanzar=1

MinSkill=87
ManaRequerido=650
StaRequerido=200
ExclusivoClase=PALADIN


&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&NIGROMANTES&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO95]
Nombre=Invocar esqueletos
Desc=Traerá tres esqueletos de los escombros.
PalabrasMagicas=MoÎ côrps
HechizeroMsg=Has invocado tres esqueletos.

Tipo=4
WAV=93
Particle=69

Invoca=1
NumNpc=503
Cant=3
MinSkill=20
ManaRequerido=150
StaRequerido=20
ExclusivoClase=NIGROMANTE
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO96]
Nombre=Invocar zombies
Desc=Invoca la ayuda del los muertos, tres zombies acudirán en tu ayuda.
PalabrasMagicas=MoÎ cámus
HechizeroMsg=Has invocado tres zombies.

Tipo=4
WAV=95

Invoca=1
NumNpc=546
Particle=69
Cant=3
MinSkill=30
ManaRequerido=250
StaRequerido=40
ExclusivoClase=NIGROMANTE
Target=4

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO97]
Nombre=Bomba mágica
Desc=Causa 20 a 30 puntos de daño en un área de dos tiles.
PalabrasMagicas=Drac né Xar
HechizeroMsg=Has lanzado Bomba mágica sobre
TargetMsg=lanzó Bomba mágica sobre vos.
PropioMsg=Has lanzado Bomba mágica sobre ti.
Tipo=1
WAV=138
FXgrh=50
loops=0


SubeHP=2
MinHP=20
MaxHP=30

MinSkill=38
ManaRequerido=50
StaRequerido=20

Target=4
Afecta=3
HechizoDeArea=1
AreaEfecto=2

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&



[HECHIZO98]
Nombre=Adivinacion
Desc=Detecta todo usuario invisible en el rango de visión de quien lo conjure.
HechizeroMsg=Has lanzado Adivinacion.
PalabrasMagicas=Vothaler Visibli

Target=4
Tipo=12
WAV=130
Particle=82
loops=1

MinSkill=100
ManaRequerido=1400
ExclusivoClase=MAGO
StaRequerido=150

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO92]
Nombre=Misil mágico
Desc=Causa 25 a 35 puntos de daño a la victima.
PalabrasMagicas=INX TO RÂ
HechizeroMsg=Has lanzado misil mágico sobre
TargetMsg=lanzo misil mágico sobre vos.
PropioMsg=Has lanzado misil mágico sobre ti.
Tipo=1
WAV=241
FXgrh=10
loops=0

SubeHP=2
MinHP=25
MaxHP=35

MinSkill=30
ManaRequerido=40
StaRequerido=20

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO93]
Nombre=Descarga eléctrica
Desc=Causa 50 a 70 puntos de daño a la victima.
PalabrasMagicas=THY KOOOL
HechizeroMsg=Has lanzado Descarga eléctrica sobre
TargetMsg=lanzo Descarga eléctrica vos.
PropioMsg=Has lanzado Descarga eléctrica sobre ti.
Tipo=1
WAV=234
Particle=49
loops=0
Resis=1

SubeHP=2
MinHP=50
MaxHP=70
    
MinSkill=65
ManaRequerido=250
StaRequerido=60

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO94]
Nombre=Ráfaga ígnea
Desc=Causa 65 a 90 puntos de daño a la victima.
PalabrasMagicas=IGNÎS XAR
HechizeroMsg=Has lanzado Ráfaga ígnea sobre
TargetMsg=lanzo Ráfaga ígnea vos.
PropioMsg=Has lanzado Ráfaga ígnea sobre ti.
Tipo=1
WAV=242
Particle=104
loops=0
Resis=1

SubeHP=2
MinHP=65
MaxHP=90
    
MinSkill=75
ManaRequerido=700
StaRequerido=100

Target=3

&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

[HECHIZO103]
Nombre=Ira de Nod
Desc=Causa 300 a 900 puntos de daño a la victima.
PalabrasMagicas=LÄTINUS MESTERMÎS
HechizeroMsg=Has lanzado la Ira de Nod sobre
TargetMsg=lanzo la Ira de Nod sobre vos.
PropioMsg=Has lanzado la Ira de Nod sobre ti.
Tipo=1
WAV=242
Particle=70
loops=0
Resis=1

SubeHP=2
MinHP=300
MaxHP=900
    
MinSkill=75
ManaRequerido=5
StaRequerido=10

Target=3');