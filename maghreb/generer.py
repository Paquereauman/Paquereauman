from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
P=Presentation(); P.slide_width=Inches(13.333); P.slide_height=Inches(7.5)
B=P.slide_layouts[6]
SABLE=RGBColor(0xF6,0xE7,0xC8); TERRE=RGBColor(0xB8,0x4A,0x1E); BLEU=RGBColor(0x1F,0x4E,0x79)
VERT=RGBColor(0x1E,0x7B,0x4A); OR=RGBColor(0xE0,0xA1,0x1B); BLANC=RGBColor(255,255,255); NOIR=RGBColor(0x2B,0x2B,0x2B)
ROUGE=RGBColor(0xC0,0x39,0x2B)
def rect(s,x,y,w,h,col,shape=MSO_SHAPE.RECTANGLE):
    r=s.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h)); r.fill.solid(); r.fill.fore_color.rgb=col; r.line.fill.background(); return r
def txt(s,x,y,w,h,t,size=20,col=NOIR,bold=False,align=PP_ALIGN.LEFT):
    tb=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame; tf.word_wrap=True
    for i,l in enumerate(t if isinstance(t,list) else [t]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.text=l; p.alignment=align
        p.font.size=Pt(size); p.font.color.rgb=col; p.font.bold=bold; p.space_after=Pt(6)
    return tb
def btn(s,x,y,w,h,t,col,size=18):
    r=rect(s,x,y,w,h,col,MSO_SHAPE.ROUNDED_RECTANGLE); tf=r.text_frame; tf.word_wrap=True
    tf.paragraphs[0].text=t; tf.paragraphs[0].font.size=Pt(size); tf.paragraphs[0].font.bold=True
    tf.paragraphs[0].font.color.rgb=BLANC; tf.paragraphs[0].alignment=PP_ALIGN.CENTER; return r
links=[]  # (shape, key)
keys={}
def slide(title,timing=None,bg=SABLE,key=None,notes=""):
    s=P.slides.add_slide(B); rect(s,0,0,13.333,7.5,bg); rect(s,0,0,13.333,1.1,TERRE)
    # frise zellige
    for i in range(27): rect(s,i*0.5,1.1,0.5,0.12,[OR,BLEU,VERT][i%3])
    txt(s,0.5,0.15,10.5,0.9,title,32,BLANC,True)
    if timing: btn(s,11.0,0.25,2.0,0.6,"⏱ "+timing,BLEU,16)
    h=btn(s,12.35,6.85,0.85,0.5,"🏠",VERT,16); links.append((h,"menu"))
    if key: keys[key]=s
    if notes: s.notes_slide.notes_text_frame.text=notes
    return s
def bullets(s,lines,y=1.5,size=22,x=0.7,w=12):
    txt(s,x,y,w,5.3,["• "+l for l in lines],size)

# 1 Titre
s=P.slides.add_slide(B); rect(s,0,0,13.333,7.5,BLEU)
for i in range(27): rect(s,i*0.5,6.9,0.5,0.6,[OR,TERRE,VERT][i%3]); rect(s,i*0.5,0,0.5,0.25,[VERT,OR,TERRE][i%3])
txt(s,0.5,1.6,12.3,1.5,"🐪  Le MAGHREB  🌴",60,OR,True,PP_ALIGN.CENTER)
txt(s,0.5,3.1,12.3,1,"Voyage ludique et interactif au pays du soleil couchant",28,BLANC,False,PP_ALIGN.CENTER)
txt(s,0.5,4.1,12.3,1,"Algérie • Maroc • Tunisie  (+ Libye & Mauritanie : le « Grand Maghreb »)",20,SABLE,False,PP_ALIGN.CENTER)
b=btn(s,5.2,5.2,2.9,0.9,"▶ COMMENCER",TERRE,24); links.append((b,"guessIntro"))
s.notes_slide.notes_text_frame.text="Séance d'environ 3 h. Cliquez sur les boutons pour naviguer (mode diaporama). Le bouton 🏠 revient au menu."

# 1b Devine où c'est
import os,json
G="/home/user/Paquereauman/maghreb/img/guess/"
CR=json.load(open(G+"credits.json")) if os.path.exists(G+"credits.json") else {}
guess=[("couscous","Qu'est-ce que c'est ? 🍽️",["Paella","Couscous","Risotto"],1,"Le couscous : plat à base de semoule de blé, inscrit au patrimoine de l'UNESCO (2020)."),
("merzouga","Où sommes-nous ? 🐪",["Sahara (Maroc)","Plage de Bretagne","Désert d'Atacama (Chili)"],0,"Les dunes de l'Erg Chebbi, près de Merzouga : le Sahara marocain."),
("chefchaouen","Quelle est cette ville bleue ? 💙",["Santorin (Grèce)","Chefchaouen (Maroc)","Venise (Italie)"],1,"Chefchaouen, la « perle bleue » du Rif, au nord du Maroc."),
("sidibou","Blanc et bleu… où est-ce ? 🏛️",["Sidi Bou Saïd (Tunisie)","Nice (France)","Lisbonne (Portugal)"],0,"Sidi Bou Saïd, village blanc et bleu près de Tunis."),
("jemaa","Quelle est cette place célèbre ? 🎪",["Place Jemaa el-Fna (Marrakech)","Place Rouge (Moscou)","Times Square (New York)"],0,"Jemaa el-Fna, à Marrakech : conteurs, musiciens et marchés le soir."),
("tajine","Comment s'appelle ce plat ? 🍲",["Fondue","Tajine","Wok"],1,"Le tajine : plat et cuisson à l'étouffée dans un pot en terre."),
("casbah","Quelle est cette vieille ville ? 🏘️",["La Casbah d'Alger (Algérie)","Montmartre (Paris)","Le Caire (Égypte)"],0,"La Casbah d'Alger, inscrite au patrimoine mondial de l'UNESCO."),
("hoggar","Quelles sont ces montagnes ? ⛰️",["Les Alpes","Le Hoggar (Algérie)","Les Rocheuses"],1,"Le Hoggar, massif volcanique du Sahara algérien."),
("oasis","Qu'est-ce que c'est ? 🌴",["Une oasis","Une forêt tropicale","Un jardin parisien"],0,"Une oasis : îlot de verdure grâce à l'eau, avec des palmiers-dattiers."),
("toubkal","Quel est ce sommet enneigé ? 🏔️",["Mont Blanc","Kilimandjaro","Toubkal (Haut Atlas, Maroc)"],2,"Le Toubkal (4 167 m), plus haut sommet d'Afrique du Nord."),
("the","Quelle boisson ? 🫖",["Thé à la menthe","Café au lait","Chocolat chaud"],0,"Le thé à la menthe, symbole d'hospitalité au Maghreb."),
("kairouan","Quel est ce monument ? 🕌",["Grande Mosquée de Kairouan (Tunisie)","Notre-Dame de Paris","Taj Mahal"],0,"La Grande Mosquée de Kairouan, l'une des plus anciennes d'Afrique du Nord (VIIe–IXe s.).")]
guess=[("@monde-musulman.png","Où se trouve le Maghreb sur cette carte ? 🗺️",["Encadré rouge, au nord-ouest de l'Afrique","Au Moyen-Orient","En Europe du Sud"],0,"Le Maghreb est au nord-ouest de l'Afrique, face à l'Europe : seulement 14 km entre le Maroc et l'Espagne (détroit de Gibraltar) !"),
("@moyen-atlas.png","Où sommes-nous ? ⛰️",["Les Alpes suisses","Le Moyen Atlas (Maroc)","Les Pyrénées"],1,"Le Moyen Atlas, près de Fès : montagnes vertes et villages.")]+guess+[
("zellige","Quel est cet art ? 🎨",["Le zellige (mosaïque)","Le vitrail","Le graffiti"],0,"Le zellige : mosaïque de petits carreaux d'argile émaillés, typique de l'art maghrébin (motifs géométriques, pas de figures humaines)."),
("@petrole.png","Que voit-on ? 🔥",["Une usine en France","Un gisement de pétrole au Sahara","Un volcan"],1,"Hassi Messaoud (Algérie) : le pétrole et le gaz sont la 1re richesse de l'Algérie."),
("@tourisme.png","Où est cet hôtel ? 🏖️",["Monastir (Tunisie)","Miami (USA)","Cannes (France)"],0,"Monastir : le tourisme balnéaire est essentiel pour la Tunisie depuis les années 1960.")]
def gpath(k): return IMGD+k[1:] if k.startswith("@") else G+k+".jpg"
IMGD="/home/user/Paquereauman/maghreb/img/"

s=slide("🔎 Jeu d'ouverture : Devine où c'est ?","15'",bg=BLEU,key="guessIntro",notes="Les équipes répondent sur ardoise avant de cliquer. 1 bonne réponse = 1 point.")
txt(s,0.6,2.0,12,1.5,["📸 Des photos vont apparaître…","Devine : où est-ce ? qu'est-ce que c'est ?","1 bonne réponse = 1 point pour ton équipe !"],30,BLANC,True,PP_ALIGN.CENTER)
b=btn(s,5.2,5.2,2.9,0.9,"C'est parti ! ➡️",TERRE,22); links.append((b,"g0"))
def cover(sl,path):
    from PIL import Image
    w,h=Image.open(path).size; W,H=13.333,5.0
    pic=sl.shapes.add_picture(path,Inches(0),Inches(1.22),Inches(W),Inches(H))
    if w/h>W/H: c=(1-(W/H)/(w/h))/2; pic.crop_left=c; pic.crop_right=c
    else: c=(1-(w/h)/(W/H))/2; pic.crop_top=c; pic.crop_bottom=c
for i,(k,q,opts,ok,fact) in enumerate(guess):
    s=slide(f"🔎 Photo {i+1}/{len(guess)} : {q}",key=f"g{i}")
    if os.path.exists(gpath(k)): cover(s,gpath(k))
    else:
        rect(s,3.2,1.6,7,4.4,RGBColor(0xE8,0xD9,0xB5),MSO_SHAPE.ROUNDED_RECTANGLE); txt(s,3.2,3.3,7,1,"📷 Insère ici la photo : "+opts[ok],20,TERRE,True,PP_ALIGN.CENTER)
    for j,o in enumerate(opts):
        b=btn(s,0.7+j*4.1,6.25,3.8,0.75,"ABC"[j]+". "+o,[TERRE,VERT,BLEU][j],16); links.append((b,f"g{i}r"))
    r=slide("✅ Réponse : "+opts[ok],bg=RGBColor(0xD5,0xF0,0xDC),key=f"g{i}r")
    if os.path.exists(gpath(k)):
        pic=r.shapes.add_picture(gpath(k),Inches(0.6),Inches(1.5),height=Inches(4.6))
        if pic.width>Inches(7): rr=Inches(7)/pic.width; pic.width=Inches(7); pic.height=int(pic.height*rr)
    txt(r,7.9,1.6,5.0,3.5,[opts[ok],"",fact],20,NOIR,False)
    c=CR.get(k)
    import re
    if c:
        txt(r,0.6,6.3,9,0.5,"Photo : Wikimedia Commons — "+re.sub("<[^>]+>","",c.get("author",""))[:60]+" — "+c.get("lic",""),10,BLEU)
    nx=btn(r,9.4,5.6,2.8,0.7,"Suivant ➡️" if i<len(guess)-1 else "Menu 🗺️",BLEU,16); links.append((nx,f"g{i+1}" if i<len(guess)-1 else "menu"))
# 2 Menu
s=slide("🗺️ Menu du voyage — clique sur une étape !",key="menu")
etapes=[("🔎 Devine où c'est ?","guessIntro","15'"),("Brise-glace","bris","10'"),("I. Carte d'identité & relief","geo","30'"),("🎬 C.O. Vidéo","video","20'"),
("Quiz géo","quiz","10'"),("📜 Histoire & données","hist","15'"),("II. Littoral & intérieur","litt","30'"),("☕ Pause","pause","10'"),
("III. Démographie & émigration","demo","30'"),("IV. Islam & conflits","islam","30'"),("📖 C.E. Yasmine","ce","20'"),("Cuisine & culture","cult","10'"),("Jeux de mots","jeux","10'"),
("Défi final","defi","5'"),("Bilan","bilan","5'")]
cols=[TERRE,BLEU,VERT,OR]
for i,(t,k,d) in enumerate(etapes):
    x=0.5+(i%4)*3.1; y=1.5+(i//4)*1.3
    b=btn(s,x,y,2.95,1.1,f"{t}  ({d})",cols[i%4],16); links.append((b,k))
txt(s,0.6,6.85,11,0.5,"Durée totale ≈ 4 h (3 séances)  •  Score des équipes au tableau  🏆",16,BLEU,True)

# 3 Brise-glace
s=slide("1. Brise-glace : le Maghreb en 1 mot","15'",key="bris",notes="Chacun dit un mot qui lui évoque le Maghreb ; on construit un nuage de mots au tableau. Puis formation de 3 équipes : Atlas, Sahara, Médina.")
bullets(s,["Tour de table : 1 mot, 1 image, 1 souvenir associé au Maghreb","On construit ensemble un nuage de mots au tableau ☁️",
"Formez 3 équipes : 🏔️ ATLAS  •  🏜️ SAHARA  •  🕌 MÉDINA","Chaque bonne réponse = 1 point — le score suit toute la séance !"])
txt(s,0.7,5.4,12,1,"« Maghreb » (المغرب) signifie en arabe : « le couchant », l'Occident — là où le soleil se couche.",20,TERRE,True)

# 4 géo
s=slide("I – Carte d'identité du Maghreb","25'",key="geo",notes="Faire lire les chiffres à voix haute, comparer avec la France (551 695 km², 68 M hab.). Valeurs arrondies : vérifier les dernières données Banque mondiale si besoin.")
rect(s,0.6,1.5,6.2,5.2,BLANC,MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s,0.8,1.6,5.9,0.6,"1) Carte d'identité régionale",22,TERRE,True)
txt(s,0.8,2.2,5.9,4.4,["Superficie : ≈ 5,8 millions km² (≈ 10 × la France métropolitaine)","Population : ≈ 107 millions d'habitants","Densité : ≈ 18 hab/km² (très inégale : côtes peuplées, Sahara presque vide)","PIB : ≈ 490 milliards $ (cumulé)","Taux de chômage : ≈ 10 à 19 % selon les pays","Statut : 5 États indépendants réunis dans l'Union du Maghreb arabe (UMA, 1989)","Source : Banque mondiale 2023 (arrondis)"],16)
rect(s,7.1,1.5,5.6,5.2,BLEU,MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s,7.3,1.6,5.3,0.6,"2) Ce qu'il faut retenir",22,OR,True)
txt(s,7.3,2.2,5.3,4.4,["🏔️ Relief varié : montagnes de l'Atlas au nord, plaines et côtes méditerranéennes, Sahara au sud","☀️ Climat : méditerranéen au nord (étés chauds, hivers doux), désertique au sud","🎶 Culture riche : musique (raï, gnaoua, chaâbi), cuisine épicée, hospitalité","🌍 Culture maghrébine : un mélange d'origines berbère (amazighe), arabe, africaine, andalouse et européenne"],16,BLANC)
s=slide("I – Carte d'identité : pays par pays","",notes="Activité : chaque équipe calcule la densité d'un pays (population ÷ superficie) et vérifie.")
hdr=["Pays","Superficie","Population","Densité","PIB","Chômage","Statut"]
data=[("🇲🇦 Maroc","446 550 km²","≈ 37 M","≈ 83","≈ 141 Md$","≈ 13 %","Monarchie const."),
("🇩🇿 Algérie","2 381 741 km²","≈ 46 M","≈ 19","≈ 240 Md$","≈ 12 %","République"),
("🇹🇳 Tunisie","163 610 km²","≈ 12 M","≈ 75","≈ 48 Md$","≈ 16 %","République"),
("🇱🇾 Libye","1 759 540 km²","≈ 7 M","≈ 4","≈ 50 Md$","≈ 19 %","République (transition)"),
("🇲🇷 Mauritanie","1 030 700 km²","≈ 5 M","≈ 5","≈ 10 Md$","≈ 10 %","Rép. islamique")]
W=[2.0,1.9,1.4,1.2,1.5,1.3,2.8]
for r,row in enumerate([hdr]+data):
    x=0.6
    for c,v in enumerate(row):
        rect(s,x,1.5+r*0.85,W[c]-0.05,0.8,TERRE if r==0 else (BLANC if r%2 else RGBColor(0xFB,0xF3,0xE2)))
        txt(s,x+0.05,1.6+r*0.85,W[c]-0.1,0.7,v,15,BLANC if r==0 else NOIR,r==0 or c==0)
        x+=W[c]
txt(s,0.6,6.65,11,0.4,"Densité en hab/km² • Source : Banque mondiale 2023 (valeurs arrondies) — capitales : Rabat, Alger, Tunis, Tripoli, Nouakchott",13,BLEU)
s=slide("I – Un monde arabo-musulman • A/ Trois États méditerranéens","",notes="Les élèves complètent leur carte au fil du cours. Distribuer la carte muette.")
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/carte-muette.png",Inches(0.6),Inches(1.5),height=Inches(5.2))
txt(s,8.0,1.5,4.8,5,["✏️ Sur ta carte muette, place :","🌊 l'océan Atlantique","🌊 la mer Méditerranée","🗺️ le Maroc, l'Algérie, la Tunisie","🇪🇸 l'Espagne","↔️ le détroit de Gibraltar","☀️ le tropique du Cancer"],20)
b=btn(s,9.4,6.0,2.8,0.7,"Correction ➡️",TERRE,16); links.append((b,"relief"))
s=slide("Le relief et les précipitations","",key="relief",notes="Questionnement oral à partir de la carte. Cliquer sur « Réponses » pour révéler.")
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/relief.png",Inches(0.5),Inches(1.4),height=Inches(5.5))
txt(s,7.5,1.5,5.4,4.5,["❓ Quel massif montagneux domine au Maghreb ?","❓ Est-il d'un seul bloc ?","❓ Quelles sont les différentes parties de l'Atlas ?","❓ Où pleut-il le plus ? Pourquoi ?","❓ Que trouve-t-on au sud ?"],19,BLEU,True)
b=btn(s,9.4,6.0,2.8,0.7,"Réponses ➡️",VERT,16); links.append((b,"reliefR"))
s=slide("Réponses : l'Atlas, un massif en plusieurs parties","",key="reliefR")
txt(s,0.6,1.5,6.3,5.3,["✅ L'ATLAS domine le Maghreb.","✅ Non, il n'est pas d'un seul bloc : il y a plusieurs parties :","   • le Rif (nord du Maroc)","   • le Moyen Atlas (près de Fès) ➡️","   • le Haut Atlas (Maroc, Toubkal 4 167 m)","   • l'Anti-Atlas","   • l'Atlas tellien et l'Atlas saharien (Algérie)","✅ Il pleut plus au nord (> 300 mm) : la montagne arrête les nuages venus de la mer.","✅ Au sud : le Sahara, des oasis et le Hoggar."],17)
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/moyen-atlas.png",Inches(7.1),Inches(1.8),width=Inches(5.8))
txt(s,7.1,4.8,5.8,0.6,"Le Moyen Atlas, près de Fès (Maroc)",16,TERRE,True,PP_ALIGN.CENTER)
s=slide("Les autres paysages du Maghreb","",notes="Questionnement oral, puis « Réponses ».")
txt(s,0.6,1.6,12,4.5,["❓ Quels sont les autres massifs montagneux du Maghreb ?","❓ Par quoi se caractérise le littoral ?","❓ Qu'est-ce qui domine au Maghreb ?","❓ Comment se nomment les îlots de verdure dans le désert ?"],24,BLEU,True)
b=btn(s,9.4,6.0,2.8,0.7,"Réponses ➡️",VERT,16); links.append((b,"oasisR"))
s=slide("Réponses : désert et oasis","",key="oasisR",notes="Insérer ici la photo de l'oasis de Bou Thramar (Insertion > Image) à la place du cadre.")
txt(s,0.6,1.5,6.3,5.3,["✅ Le Hoggar (sud de l'Algérie).","✅ Le littoral se caractérise par des plaines.","✅ Le désert (Sahara) domine, avec des îlots de verdure.","✅ Ces îlots s'appellent des OASIS 🌴"],20)
r=rect(s,7.1,1.6,5.8,3.4,RGBColor(0xE8,0xD9,0xB5),MSO_SHAPE.ROUNDED_RECTANGLE)
txt(s,7.1,2.9,5.8,0.8,"📷 Photo : l'oasis de Bou Thramar",18,TERRE,True,PP_ALIGN.CENTER)
txt(s,7.1,5.1,5.8,0.6,"L'oasis de Bou Thramar (Maroc)",16,TERRE,True,PP_ALIGN.CENTER)
s=slide("Doc. 1 p. 222 : les précipitations","",notes="Faire repérer les lignes 300 mm et 150 mm sur la carte.")
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/relief.png",Inches(0.5),Inches(1.4),height=Inches(5.5))
txt(s,7.5,1.5,5.4,3,["❓ Quel est le niveau des précipitations dans la région ?","❓ Où pleut-il le plus ?"],20,BLEU,True)
b=btn(s,7.5,3.6,2.6,0.7,"Réponse 1",VERT,16); links.append((b,"pluieR"))
s=slide("Réponse : des précipitations globalement faibles","",key="pluieR")
txt(s,0.6,1.6,12,4.5,["✅ Les précipitations sont globalement FAIBLES.","• Plus de 300 mm/an seulement sur une étroite bande au nord (littoral, Atlas).","• Moins de 150 mm/an dans le Sahara : un climat désertique.","💡 Comparaison : Paris reçoit environ 640 mm de pluie par an."],22)
s=slide("Précipitations et climats","",notes="Préciser qu'en dessous de 250 mm de précipitations, il n'y a plus d'agriculture possible (sans irrigation).")
txt(s,0.6,1.6,12,4,["❓ Cette faiblesse des précipitations est-elle généralisée ?","❓ Lesquelles des régions sont plus arrosées ?","❓ Quels sont les climats du Maghreb ?"],24,BLEU,True)
b=btn(s,9.4,6.0,2.8,0.7,"Réponses ➡️",VERT,16); links.append((b,"climR"))
s=slide("Réponses : deux climats","",key="climR")
txt(s,0.6,1.5,6.3,5.3,["✅ Non : certaines régions sont un peu plus arrosées.","✅ Les plaines littorales et les massifs montagneux.","⚠️ En dessous de 250 mm/an, plus d'agriculture possible !","✅ Au Nord : climat MÉDITERRANÉEN 🌿","✅ Au Sud : climat DÉSERTIQUE 🏜️","✏️ Sur ta carte : trace l'isohyète des 250 mm et complète la légende."],19)
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/isohyete.png",Inches(7.1),Inches(1.5),width=Inches(5.8))
txt(s,7.1,6.4,5.8,0.5,"━ isohyète 250 mm (ligne bleue)",15,BLEU,True,PP_ALIGN.CENTER)
s=slide("✏️ Complète le tableau","",notes="Les élèves complètent le tableau sur leur fiche, puis correction.")
rect(s,2.0,1.6,9.3,0.9,TERRE); txt(s,2.0,1.75,9.3,0.7,"Principaux caractères du Maghreb",24,BLANC,True,PP_ALIGN.CENTER)
for i in range(3):
    rect(s,2.0,2.6+i*1.2,9.3,1.1,BLANC); txt(s,2.0,2.85+i*1.2,9.3,0.7,"1… 2… 3… ?"[0:0]+["1. ……………………","2. ……………………","3. ……………………"][i],22,NOIR,False,PP_ALIGN.CENTER)
b=btn(s,9.4,6.3,2.8,0.6,"Correction ➡️",VERT,16); links.append((b,"tabR"))
s=slide("Correction du tableau","",key="tabR")
rect(s,2.0,1.6,9.3,0.9,TERRE); txt(s,2.0,1.75,9.3,0.7,"Principaux caractères du Maghreb",24,BLANC,True,PP_ALIGN.CENTER)
for i,t in enumerate(["Importance du désert","Présence d'un massif montagneux : l'Atlas","Les domaines bioclimatiques méditerranéen et désertique, marqués par l'aridité"]):
    rect(s,2.0,2.6+i*1.2,9.3,1.1,BLANC); txt(s,2.1,2.75+i*1.2,9.1,0.9,t,20,NOIR,True,PP_ALIGN.CENTER)
s=slide("B/ Le Maghreb appartient au monde arabo-musulman","",notes="Présentation : http://fr.calameo.com/read/0000054193e34050e91d9")
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/monde-musulman.png",Inches(0.5),Inches(1.5),height=Inches(5.2))
txt(s,6.9,1.6,6,4,["❓ D'après ce document et ce que tu as vu sur l'Afrique, quelle est la religion dominante au Maghreb ?","❓ Depuis quand ?"],21,BLEU,True)
b=btn(s,9.4,6.0,2.8,0.7,"Réponses ➡️",VERT,16); links.append((b,"isR"))
s=slide("Réponses : le Maghreb musulman","",key="isR")
txt(s,0.6,1.8,12,4,["✅ La religion musulmane (islam) domine.","✅ Depuis les VIIe–VIIIe siècles, avec les conquêtes arabo-musulmanes.","🕋 Lieux saints de l'islam : La Mecque et Médine (Arabie saoudite)."],24)
s=slide("📖 C.E. — Doc. 4 p. 221 : Les Berbères","15'",notes="Lecture silencieuse 5', questions en binômes, correction collective.")
txt(s,0.6,1.4,12.1,5.5,["« Les Berbères habitent en Algérie, dans les montagnes de Kabylie et des Aurès, et surtout au Maroc, dans les montagnes du Rif et dans l'Atlas.",
"Les Berbères sont au moins bilingues. Ils parlent le berbère et l'arabe. Seuls les très jeunes enfants, les femmes ou les vieillards n'ayant pas séjourné en ville ne parlent que le berbère : dans la ville maghrébine, en effet, on parle arabe, à l'exception de quelques villes comme Agadir au Maroc.",
"Tous les Maghrébins, qu'ils soient arabes ou berbères, accordent une grande valeur sentimentale à des traits de la culture berbère. Il s'agit notamment du « burnous », ce long manteau de laine à capuchon pour les hommes, et surtout du « couscous »¹, le plat le plus prestigieux de la cuisine maghrébine. »",
"D'après Yves Lacoste, Maghreb, 1995.   ¹ Plat à base de semoule de blé."],17)
b=btn(s,10.0,6.85,2.2,0.5,"Questions ➡️",VERT,14); links.append((b,"berQ"))
s=slide("C.E. Les Berbères — Questions","",key="berQ")
bullets(s,["1. Quels sont les deux peuples présents au Maghreb ?","2. Où habitent les Berbères ? Cite 2 pays et 3 régions.","3. Vrai/Faux : les Berbères ne parlent que le berbère. Justifie avec le texte.","4. Quelle langue parle-t-on dans les villes ? Quelle exception ?","5. Qu'est-ce que le « burnous » ?","6. Quel est « le plat le plus prestigieux » ? De quoi est-il fait ?","7. Qui est l'auteur ? En quelle année ?"],1.5,20)
b=btn(s,9.8,6.1,2.4,0.7,"Correction",TERRE,16); links.append((b,"berC"))
s=slide("C.E. Les Berbères — Correction","",key="berC")
bullets(s,["1. Les Arabes et les Berbères.","2. En Algérie (Kabylie, Aurès) et surtout au Maroc (Rif, Atlas).","3. FAUX : « Les Berbères sont au moins bilingues » (berbère et arabe).","4. L'arabe, sauf quelques villes comme Agadir (Maroc).","5. Un long manteau de laine à capuchon pour les hommes.","6. Le couscous, à base de semoule de blé.","7. Yves Lacoste, 1995."],1.5,20)
s=slide("Doc. 5 p. 221 : les Arabes et les Berbères","",notes="Présentation Calaméo : https://www.calameo.com/read/000005419914424b872ed")
s.shapes.add_picture("/home/user/Paquereauman/maghreb/img/arabes-berberes.png",Inches(0.5),Inches(1.4),height=Inches(5.5))
txt(s,7.6,1.5,5.3,4.5,["❓ Où vivent les arabophones ?","❓ Où vivent les berbérophones ?","❓ Quelle part de la population parle berbère au Maroc ? en Algérie ?","❓ Pourquoi le Sahara est-il presque vide ?"],19,BLEU,True)
b=btn(s,9.4,6.0,2.8,0.7,"Réponses ➡️",VERT,16); links.append((b,"abR"))
s=slide("Réponses : Arabes et Berbères","",key="abR")
bullets(s,["Arabophones : plaines littorales et villes (Casablanca, Rabat, Alger, Oran, Tunis).","Berbérophones : surtout dans les montagnes (Rif, Atlas, Kabylie, Aurès) et autour d'Agadir.","Berbérophones : environ 33 % au Maroc et 25 % en Algérie (très peu en Tunisie).","Le Sahara : moins de 10 hab/km², car le climat désertique y rend la vie difficile.","➡️ Les Berbères sont le peuple d'origine ; les Arabes sont arrivés avec l'islam (VIIe s.)."],1.6,20)
IMG="/home/user/Paquereauman/maghreb/img/"
TR=RGBColor(0x1A,0x2F,0xC8)
n=[0]
def doc(title,img,qs,ans,trace=None,timing=None,key=None,notes="",wide=False):
    n[0]+=1; k=f"doc{n[0]}"
    s=slide(title,timing,key=key,notes=notes)
    pic=s.shapes.add_picture(IMG+img,Inches(0.5),Inches(1.4),height=Inches(5.4))
    if pic.width>Inches(7.2): 
        r=Inches(7.2)/pic.width; pic.width=Inches(7.2); pic.height=int(pic.height*r)
    x=0.5+pic.width/914400+0.3
    txt(s,x,1.5,13.0-x,4.6,["❓ "+q for q in qs],18,BLEU,True)
    bt=btn(s,10.2,6.15,2.2,0.6,"Réponses ➡️",VERT,15); links.append((bt,k))
    r=slide("Réponses — "+title.split(":")[-1].strip(),key=k)
    txt(r,0.6,1.5,12.1,3.8 if trace else 5.2,["✅ "+a for a in ans],19)
    if trace:
        rect(r,0.6,5.25,12.1,1.5,RGBColor(0xE6,0xEA,0xFB),MSO_SHAPE.ROUNDED_RECTANGLE)
        txt(r,0.8,5.3,11.7,1.45,["✏️ Trace écrite : "+trace],17,TR,True)
def trace(title,lines):
    s=slide(title); rect(s,0.6,2.0,12.1,3.2,RGBColor(0xE6,0xEA,0xFB),MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s,0.9,2.2,11.5,3,["✏️ Trace écrite :"]+lines,22,TR,True)
trace("Trace écrite — B/ Le monde arabo-musulman",["La population du Maghreb est composée majoritairement d'Arabes, mais les Berbères sont également nombreux.","La religion de la majorité de la population est l'islam."])

# ===== II =====
s=slide("II. Un littoral plus attractif que l'intérieur","30'",bg=BLEU,key="litt",notes="Présentation : http://fr.calameo.com/read/00000541937ce663e2326")
txt(s,0.6,2.3,12,1,"Séance 2 : littoral et intérieur",36,OR,True,PP_ALIGN.CENTER)
txt(s,0.6,3.5,12,1,"Problématique : en quoi le littoral est-il le cœur économique du Maghreb ?",26,BLANC,False,PP_ALIGN.CENTER)
txt(s,0.6,4.8,12,1,"A/ La répartition de la population   •   B/ Un littoral plus développé",20,SABLE,False,PP_ALIGN.CENTER)
doc("A/ Doc. 4 p. 223 : la répartition de la population","repartition.png",
 ["D'après cette carte, où se concentre la population du Maghreb ?","Quelles sont les densités à l'intérieur du territoire ?","Quelles sont les causes de cette répartition ?","Où se situent les villes ? Quelles sont les plus importantes ?"],
 ["Sur le littoral.","Très faibles (moins de 10 hab/km²).","Le désert fait fuir les populations.","Sur le littoral : Casablanca, Rabat, Alger ; puis Agadir, Tanger, Oran, Tunis.","✏️ Sur ta carte : place les zones de fortes densités et les villes, complète la légende."],
 "Le territoire du Maghreb est principalement occupé par le désert, très faiblement peuplé. Au contraire, le Maghreb littoral regroupe 80 % de la population.")
doc("B/ Doc. 3 p. 227 : un gisement pétrolier au Sahara (Hassi Messaoud)","petrole.png",
 ["Où a été prise la photographie ?","Décris ce que tu vois.","Que produit-on ?"],["Au Sahara (Algérie).","Un gisement : torchères, réservoirs, fumées.","Du pétrole."])
doc("Doc. 4 p. 227 : les ressources intérieures","ressources.png",
 ["Quels sont les grands gisements de pétrole ?","Quels sont les autres gisements ?","Comment appelle-t-on ces deux produits ?"],
 ["Hassi Messaoud et Edjeleh.","Les gisements de gaz, comme Hassi R'Mel.","Des HYDROCARBURES (pétrole + gaz). 📖 Copie la définition.","✏️ Place les gisements de pétrole sur ta carte et complète la légende."])
doc("Doc. 4 p. 229 : le développement du littoral","littoral.png",
 ["Que devient le pétrole extrait dans le désert ?","Où se situent les régions industrielles ?"],
 ["Il est surtout exporté vers l'Europe, et utilisé dans les centres industriels.","Sur le littoral : Casablanca, Oran-Arzew, Alger, Annaba, Tunis…"])
doc("📖 C.E. — Doc. 3 p. 229 : le complexe industriel d'Arzew","arzew.png",
 ["Où se situe Arzew ?","Quelle est la surface de la zone industrielle ?","Combien de personnes y travaillent ?","Quelles sont les activités d'Arzew ?","Comment s'appelle ce type de centre industriel ?"],
 ["Au nord de l'Algérie, sur le littoral.","1 500 hectares.","10 000 travailleurs.","Raffinage du pétrole, liquéfaction du gaz, puis exportation (port d'Arzew + Bethioua).","Un COMPLEXE INDUSTRIALO-PORTUAIRE. 📖 Copie la définition."],
 "Si le désert possède d'importantes réserves en hydrocarbures, ceux-ci sont acheminés vers les ports d'exportation et vers les complexes industrialo-portuaires.")
doc("Doc. 1 p. 228 : le conditionnement des agrumes (Maroc)","agrumes.png",
 ["Décris l'image.","Dans quel pays cela se situe-t-il ?","Quels sont les agrumes ?","Qui sont les travailleurs ?","Où se situent les régions agricoles ? Quelles autres activités ?"],
 ["Une usine de traitement des agrumes (sélection, emballage).","Au Maroc.","Des oranges.","Surtout des femmes, des ouvrières.","Dans les plaines littorales : grande céréaliculture, viticulture. ✏️ Ajoute « agriculture moderne » à la légende."],
 "Le littoral est aussi une région agricole où l'on pratique une agriculture moderne pour l'exportation.")
doc("📖 C.E. — Doc. 5 et 6 p. 229 : le tourisme en Tunisie","tourisme.png",
 ["Quand est apparu le tourisme en Tunisie ?","Qui sont les touristes ?","Quelle place a le tourisme dans l'économie ?","Qu'est-ce qui attire les touristes ? Comment viennent-ils ?","Quelles sont les grandes stations ?"],
 ["Dans les années 1960.","Les Européens, puis les Algériens et Libyens, et les Tunisiens eux-mêmes.","Une place importante (5e–6e rang méditerranéen).","Le soleil, la mer ; ils viennent en avion (aéroport nécessaire).","Hammamet et Monastir."],
 "Depuis les années 1960, le Maroc et la Tunisie se sont lancés dans l'activité touristique.")

# ===== III =====
s=slide("III. Les problèmes démographiques et sociaux","30'",bg=BLEU,key="demo",notes="Présentation : http://fr.calameo.com/read/000005419fb38f20f0ef7")
txt(s,0.6,2.3,12,1,"Séance 3",36,OR,True,PP_ALIGN.CENTER)
txt(s,0.6,3.5,12,1,"Problématique : quels sont les problèmes du Maghreb d'aujourd'hui ?",26,BLANC,False,PP_ALIGN.CENTER)
txt(s,0.6,4.8,12,1,"A/ Une population en augmentation   •   B/ Un foyer d'émigration vers l'Europe",20,SABLE,False,PP_ALIGN.CENTER)
doc("A/ Doc. 1 p. 224 : l'évolution de la population","evolution.png",
 ["Quelle est la population des trois pays en 2000 ?","Comment a-t-elle évolué depuis 1950 ?","🎲 Défi calcul : par combien la population de l'Algérie a-t-elle été multipliée ?"],
 ["Algérie : 31 M • Maroc : 30 M • Tunisie : 9,5 M.","Elle a fortement augmenté : environ × 3 en 50 ans.","9 M → 31 M : environ × 3,4.","💡 Aujourd'hui : Algérie ≈ 46 M, Maroc ≈ 37 M, Tunisie ≈ 12 M."],
 "Le Maghreb compte environ 70 millions d'habitants (en 2000). La population a fortement augmenté depuis 50 ans.")
doc("📖 C.E. — Doc. 2 p. 224 : de lourdes conséquences","consequences.png",
 ["Évolution des effectifs scolaires algériens entre 1966 et 1988 ?","Quelle part du budget de l'État va à l'éducation ?","Tout le monde est-il scolarisé ? Dans quelles conditions ?","Quelle est la 2e conséquence de la hausse du nombre de jeunes ?"],
 ["Forte hausse : de 1,5 million à plus de 6 millions d'élèves.","30 %, c'est-à-dire beaucoup.","Non, seulement 80 % des jeunes, dans de mauvaises conditions.","Le chômage : 250 000 jeunes arrivent chaque année sur le marché du travail, mais seulement 80 000 emplois créés."])
doc("Doc. 4 p. 225 : le taux d'urbanisation","urbanisation.png",
 ["Comment a évolué la population urbaine ?","À quoi ce phénomène est-il dû ?"],
 ["Forte hausse : d'environ 40 % à près de 60 % de la population (1970 → 2000).","Au chômage dans les campagnes (exode rural) et à la forte croissance de la population."])
doc("📖 C.E. — Doc. 6 p. 225 : les problèmes dans les villes","villes.png",
 ["Quels problèmes la forte urbanisation a-t-elle entraînés ?","Combien de personnes par logement en moyenne ?","Où se concentrent les mécontentements ?"],
 ["Logement (bidonvilles), circulation, transports, manque d'emplois.","8 personnes par logement.","Dans les quartiers pauvres des villes : « la poudrière du pays »."])
# schéma interactif
s=slide("✏️ Schéma : les conséquences de la croissance démographique","",notes="Cliquer sur chaque « ? » pour révéler la conséquence.")
btn(s,3.4,1.5,6.5,1.0,"Augmentation de la population",TERRE,26)
for i,(t,k) in enumerate([("Scolarisation de tous les enfants difficile","sc1"),("Exode rural, forte croissance urbaine, bidonvilles","sc2"),("Chômage","sc3")]):
    x=0.8+i*4.1
    con=s.shapes.add_connector(1,Inches(6.65),Inches(2.5),Inches(x+1.75),Inches(3.7)); con.line.color.rgb=NOIR; con.line.width=Pt(2.5)
    b=btn(s,x,3.7,3.5,1.3,f"❓ Conséquence {i+1}",BLEU,18); links.append((b,"schema"))
s=slide("✏️ Schéma complété","",key="schema")
btn(s,3.4,1.5,6.5,1.0,"Augmentation de la population",TERRE,26)
for i,t in enumerate(["Scolarisation de tous les enfants difficile","Exode rural, forte croissance urbaine, bidonvilles","Chômage"]):
    x=0.8+i*4.1
    con=s.shapes.add_connector(1,Inches(6.65),Inches(2.5),Inches(x+1.75),Inches(3.7)); con.line.color.rgb=NOIR; con.line.width=Pt(2.5)
    btn(s,x,3.7,3.5,1.3,t,VERT,18)
doc("B/ 📖 C.E. — Doc. 2 p. 230 : les migrations au Maghreb","migrations.png",
 ["Face à la pauvreté, que fait une partie de la population ?","Quelles sont les motivations ?","Vers quels pays se dirigent les émigrés ?","Ce phénomène est-il du passé ou du présent ?","Quelle part des revenus du Maroc représentait l'argent envoyé en 1993 ?"],
 ["Elle émigre vers l'Europe.","Le travail.","La France, le Benelux, l'Italie et l'Espagne.","Surtout du passé : l'Europe a fermé ses frontières ; aujourd'hui surtout regroupement familial (et clandestins).","7,5 % (4,5 % pour la Tunisie, 2,5 % pour l'Algérie)."],
 "Depuis environ 40 ans, beaucoup de Maghrébins ont choisi d'émigrer vers l'Europe pour travailler. Aujourd'hui, l'Europe a fermé ses frontières et seul le regroupement familial est autorisé. Certains tentent toutefois encore d'émigrer : on parle de clandestins.")
# ===== IV islam & conflits =====
s=slide("IV. Islam, Moyen-Orient et conflits","30'",bg=BLEU,key="islam",notes="Séance sensible : rappeler que l'immense majorité des musulmans rejette le terrorisme. Ne pas confondre islam et islamisme.")
txt(s,0.6,2.2,12,1,"Une région riche d'histoire… mais aussi de tensions",34,OR,True,PP_ALIGN.CENTER)
txt(s,0.6,3.6,12,1.5,["A/ L'arrivée de l'islam   •   B/ Sunnites et chiites","C/ Les conflits au Moyen-Orient   •   D/ Printemps arabe et terrorisme"],22,BLANC,False,PP_ALIGN.CENTER)
s=slide("A/ Quand l'islam est-il arrivé au Maghreb ?","")
ev=[("610","Mahomet reçoit la Révélation à La Mecque"),("622","Hégire : départ pour Médine (an 1 du calendrier musulman)"),("632","Mort de Mahomet"),("670","Fondation de Kairouan (Tunisie), 1re grande ville musulmane du Maghreb"),("711","Tariq ibn Ziyad traverse vers l'Espagne (Gibraltar = « Djebel Tariq »)")]
for i,(d,t) in enumerate(ev):
    btn(s,0.6,1.5+i*0.95,1.6,0.8,d,[TERRE,BLEU,VERT,OR,TERRE][i],20); txt(s,2.4,1.6+i*0.95,10.4,0.8,t,19)
txt(s,0.6,6.35,12,0.5,"➡️ Les Berbères se convertissent peu à peu ; l'arabe devient la langue de la religion et des villes.",16,TERRE,True)
doc("B/ Sunnites et chiites : quelle différence ?","monde-musulman.png",
 ["Pourquoi les musulmans se sont-ils divisés après 632 ?","Que pensent les sunnites ? les chiites ?","Le Maghreb est-il sunnite ou chiite ?","Où vivent surtout les chiites ?"],
 ["Désaccord sur la succession de Mahomet (qui doit diriger les musulmans ?).","Sunnites (≈ 85-90 %) : le chef (calife) est choisi parmi les fidèles. Chiites (≈ 10-15 %) : il doit descendre d'Ali, gendre de Mahomet (bataille de Kerbala, 680).","Presque entièrement sunnite (rite malékite).","Iran, Irak, Bahreïn, Liban (en partie)."],
 "Le Maghreb est musulman sunnite. La division sunnites/chiites, née d'une querelle de succession, alimente aujourd'hui des rivalités au Moyen-Orient (Arabie saoudite sunnite / Iran chiite).")
s=slide("C/ Le Moyen-Orient : une région de conflits","",notes="Rester factuel ; laisser les élèves poser leurs questions.")
cf=[("1948 →","Conflit israélo-palestinien"),("1980–1988","Guerre Iran–Irak"),("2003","Guerre en Irak (intervention américaine)"),("2011 →","Guerre civile en Syrie"),("2014–2019","L'État islamique (Daech) en Irak et en Syrie"),("2015 →","Guerre au Yémen (rivalité Arabie saoudite / Iran)")]
for i,(d,t) in enumerate(cf):
    btn(s,0.6,1.5+i*0.82,2.1,0.7,d,ROUGE if i%2 else TERRE,17); txt(s,2.9,1.55+i*0.82,9.8,0.7,t,19)
txt(s,0.6,6.5,12,0.5,"💡 Causes mêlées : pétrole, frontières, religion (sunnites/chiites), rivalités entre puissances.",16,BLEU,True)
s=slide("D/ Le Printemps arabe (2010–2011)","",notes="Bouazizi : vendeur ambulant de Sidi Bouzid.")
pa=[("🇹🇳 Tunisie","17 déc. 2010 : immolation de Mohamed Bouazizi → départ de Ben Ali (14 janv. 2011)"),("🇪🇬 Égypte","Février 2011 : chute de Moubarak"),("🇱🇾 Libye","2011 : guerre civile, mort de Kadhafi ; le pays reste divisé"),("🇲🇦 Maroc","2011 : manifestations, nouvelle Constitution"),("🇩🇿 Algérie","2019 : le Hirak pousse Bouteflika à partir"),("🇸🇾 Syrie","2011 : révolte réprimée → guerre civile")]
for i,(c,t) in enumerate(pa):
    txt(s,0.6,1.5+i*0.85,2.6,0.8,c,20,TERRE,True); txt(s,3.2,1.52+i*0.85,9.6,0.8,t,18)
txt(s,0.6,6.6,12,0.5,"❓ Débat : pourquoi les jeunes se sont-ils révoltés ? (chômage, corruption, manque de libertés)",16,BLEU,True)
doc("D/ Le terrorisme : de quoi parle-t-on ?","monde-musulman.png",
 ["Islam et islamisme, est-ce la même chose ?","Qu'est-ce que la « décennie noire » en Algérie ?","Cite des attentats au Maghreb.","Qui sont les principales victimes du terrorisme islamiste ?"],
 ["Non ! L'islam est une religion ; l'islamisme radical est une idéologie politique violente rejetée par la grande majorité des musulmans.","1991-2002 : guerre entre l'État et des groupes islamistes armés (GIA…), des dizaines de milliers de morts.","Casablanca (2003), musée du Bardo et Sousse en Tunisie (2015)…","Surtout des musulmans eux-mêmes, au Maghreb et au Moyen-Orient."],
 "Le Maghreb a connu le terrorisme islamiste (décennie noire en Algérie, attentats au Maroc et en Tunisie). Il ne faut pas confondre islam et terrorisme.")
s=slide("🎲 Vrai ou faux ? (islam & conflits)","")
vf=[("L'islam est arrivé au Maghreb au VIIe siècle.","VRAI"),("Le Maghreb est majoritairement chiite.","FAUX : sunnite"),("Le Printemps arabe a commencé en Tunisie.","VRAI (déc. 2010)"),("Tous les musulmans soutiennent le terrorisme.","FAUX : ils en sont les 1res victimes"),("Gibraltar vient de « Djebel Tariq ».","VRAI")]
for i,(q,a) in enumerate(vf):
    txt(s,0.6,1.5+i*0.95,8.4,0.8,f"{i+1}. {q}",19)
    b=btn(s,9.3,1.5+i*0.95,3.4,0.75,"Révéler",[TERRE,BLEU,VERT,OR,TERRE][i],15); links.append((b,"vfR"))
s=slide("Correction du vrai/faux","",key="vfR")
txt(s,0.6,1.5,12,5,[f"{i+1}. {q}  ➜  {a}" for i,(q,a) in enumerate(vf)],20)
s=slide("Paysages : de la mer au désert","")
for i,(t,d,c) in enumerate([("🌊 Méditerranée","Côtes, ports, plages au nord",BLEU),("🏔️ Atlas","Montagnes jusqu'à 4 167 m (Toubkal)",VERT),("🏜️ Sahara","Le plus grand désert chaud du monde",OR),("🌴 Oasis","Palmiers-dattiers, eau et vie",TERRE)]):
    r=rect(s,0.6+i*3.15,1.7,2.95,4.2,c,MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s,0.7+i*3.15,2.0,2.75,1,t,24,BLANC,True,PP_ALIGN.CENTER); txt(s,0.7+i*3.15,3.3,2.75,2,d,18,BLANC,False,PP_ALIGN.CENTER)
txt(s,0.6,6.1,12,0.7,"🎯 Activité : à l'oral, décris ton paysage préféré avec 3 adjectifs.",20,TERRE,True)

# vidéo
s=slide("🎬 C.O. Compréhension orale : vidéo","20'",key="video",notes="Remplacer le lien par la vidéo choisie (ou insérer un fichier .mp4 : Insertion > Vidéo). 1er visionnage sans son, 2e avec son, 3e pour vérifier.")
r=rect(s,0.8,1.6,6.6,4.2,NOIR); 
v=btn(s,2.6,3.2,3.0,1.0,"▶ LANCER LA VIDÉO",ROUGE,20)
v.click_action.hyperlink.address="https://www.youtube.com/results?search_query=le+maghreb+documentaire+maroc+alg%C3%A9rie+tunisie"
txt(s,0.8,5.9,6.6,0.8,"(Clique pour ouvrir la vidéo — ou insère ton .mp4 ici)",14,BLEU,False,PP_ALIGN.CENTER)
txt(s,7.7,1.6,5.2,5,["📺 Visionnage 1 — sans le son :","   Quels lieux, couleurs, objets vois-tu ?","📺 Visionnage 2 — avec le son :","   Note 3 pays, 2 villes, 1 plat.","📺 Visionnage 3 :","   Réponds à la fiche vidéo ➡️"],18)
b=btn(s,9.5,6.0,2.7,0.7,"Fiche vidéo ➡️",VERT,16); links.append((b,"fiche"))
s=slide("Fiche vidéo : as-tu bien regardé ?","",key="fiche")
bullets(s,["Quels pays sont présentés ?","Cite une ville et ce qu'on y voit.","Quel paysage t'a le plus surpris ? Pourquoi ?","Quelles langues entend-on ?","Vrai ou faux : le Maghreb est uniquement désertique."],1.5,22)
b=btn(s,9.8,6.0,2.4,0.7,"Correction",TERRE,16); links.append((b,"ficheC"))
s=slide("Fiche vidéo : correction","",key="ficheC")
bullets(s,["Maroc, Algérie, Tunisie (et parfois Libye, Mauritanie)","Ex. : Marrakech (place Jemaa el-Fna), Alger (la Casbah), Tunis (la médina)","Réponse libre — justifie avec « parce que… »","Arabe (darija), tamazight (berbère), français","FAUX ❌ : il y a aussi la mer, les montagnes, les villes, les oasis !"])

# quiz
quiz=[("Quelle est la capitale du Maroc ?",["Casablanca","Rabat","Marrakech"],1),
("Quel est le plus grand pays d'Afrique ?",["Algérie","Libye","Égypte"],0),
("Que signifie « Maghreb » ?",["Le désert","Le couchant","La mer"],1),
("Quelle montagne traverse le Maghreb ?",["Les Alpes","L'Atlas","L'Himalaya"],1),
("Quelle mer borde le Maghreb au nord ?",["Mer Rouge","Méditerranée","Mer Noire"],1)]
first=True
for i,(q,opts,ok) in enumerate(quiz):
    s=slide(f"Quiz géo — Question {i+1}/5","10'" if first else None,key="quiz" if first else None); first=False
    txt(s,0.6,1.6,12,1,"❓ "+q,30,BLEU,True,PP_ALIGN.CENTER)
    for j,o in enumerate(opts):
        b=btn(s,1.0+j*3.9,3.3,3.6,1.4,"ABC"[j]+". "+o,[TERRE,VERT,BLEU][j],24)
        links.append((b,f"q{i}ok" if j==ok else f"q{i}ko"))
    ok_s=slide("✅ BRAVO ! +1 point",bg=RGBColor(0xD5,0xF0,0xDC),key=f"q{i}ok")
    txt(ok_s,0.6,2.5,12,1.5,f"Bonne réponse : {opts[ok]} 🎉",40,VERT,True,PP_ALIGN.CENTER)
    nxt=btn(ok_s,5.2,4.5,2.9,0.9,"Suivant ➡️",BLEU,22); links.append((nxt,f"q{i+1}" if i<4 else "hist"))
    ko=slide("❌ Oups ! Essaie encore",bg=RGBColor(0xF8,0xD7,0xD3),key=f"q{i}ko")
    txt(ko,0.6,2.5,12,1.5,"Ce n'est pas la bonne réponse… 🤔",40,ROUGE,True,PP_ALIGN.CENTER)
    back=btn(ko,5.2,4.5,2.9,0.9,"⬅️ Réessayer",TERRE,22); links.append((back,f"q{i}"))
    keys[f"q{i}"]=P.slides[P.slides.index(ok_s)-1]

# histoire
s=slide("📜 Histoire : la frise du temps","15'",key="hist",notes="Activité : distribuer les étiquettes dans le désordre, les équipes les remettent dans l'ordre.")
ev=[("Antiquité","Berbères (Amazighs), Carthage (814 av. J.-C.), Rome"),("VIIe s.","Conquête arabe : islam et langue arabe"),("XVIe s.","Empire ottoman (Alger, Tunis, Tripoli) ; Maroc indépendant"),("1830–1912","Colonisation : France (Algérie, Tunisie, Maroc), Italie (Libye)"),("1951–1962","Indépendances")]
rect(s,0.6,3.6,12.1,0.12,TERRE)
for i,(d,t) in enumerate(ev):
    x=0.6+i*2.45; rect(s,x+0.9,3.35,0.6,0.6,OR,MSO_SHAPE.OVAL)
    txt(s,x,2.2,2.4,1,d,20,BLEU,True,PP_ALIGN.CENTER); txt(s,x,4.2,2.4,1.6,t,15,NOIR,False,PP_ALIGN.CENTER)
txt(s,0.6,6.0,12,0.8,"🗣️ Langues : arabe (darija), tamazight (ⵜⴰⵎⴰⵣⵉⵖⵜ), français • Salam = bonjour • Choukran = merci • Azul = bonjour (berbère)",16,TERRE,True)
s=slide("📜 Histoire : quelques données clés","")
rows=[("1951","Indépendance de la Libye"),("1956","Indépendance du Maroc (2 mars) et de la Tunisie (20 mars)"),("1954–1962","Guerre d'Algérie → indépendance le 5 juillet 1962"),("1960","Indépendance de la Mauritanie"),("1989","Création de l'Union du Maghreb arabe (Marrakech)"),("2011","« Printemps arabe » : départ de Ben Ali (Tunisie), chute de Kadhafi (Libye)"),("2020","Le couscous inscrit au patrimoine de l'UNESCO (candidature commune)")]
for i,(d,t) in enumerate(rows):
    btn(s,0.6,1.5+i*0.75,2.0,0.65,d,[TERRE,BLEU,VERT][i%3],18); txt(s,2.8,1.55+i*0.75,10,0.6,t,18)
txt(s,0.6,6.8,11,0.4,"🎯 Question : combien d'années séparent l'indépendance de la Libye de celle de l'Algérie ? (11 ans)",15,BLEU,True)
s=slide("☕ PAUSE — thé à la menthe !","15'",bg=RGBColor(0xE3,0xF1,0xE6),key="pause")
txt(s,0.6,2.4,12,2,["🫖 15 minutes de pause","Le thé à la menthe, symbole d'hospitalité, se sert « de haut » pour faire de la mousse !"],30,VERT,True,PP_ALIGN.CENTER)

# C.E.
s=slide("📖 C.E. Compréhension écrite : la lettre de Yasmine","20'",key="ce",notes="Lecture silencieuse 10', questions en binômes 15', correction collective 10'.")
txt(s,0.6,1.5,12.1,5.3,["Lettre de Yasmine, depuis Marrakech",
"Chère Léa,",
"Me voici enfin à Marrakech, chez ma grand-mère ! Hier matin, nous sommes allées au souk : quel spectacle ! Des épices de toutes les couleurs, des tapis, des lanternes en cuivre… On a acheté du cumin et du safran pour le couscous du vendredi.",
"Le soir, la place Jemaa el-Fna se transforme : il y a des conteurs, des musiciens gnaoua et des dizaines de petits restaurants. J'ai goûté une harira, une soupe délicieuse.",
"Samedi, nous partons trois jours dans le désert, près de Merzouga. Nous dormirons sous une tente berbère et nous monterons à dos de dromadaire pour voir le lever du soleil sur les dunes. J'ai hâte !",
"Je t'envoie plein de soleil. Bisous, Yasmine"],17)
b=btn(s,10.2,6.85,2.0,0.5,"Questions ➡️",VERT,14); links.append((b,"ceq"))
s=slide("C.E. — Questions","",key="ceq")
bullets(s,["1. Qui écrit ? À qui ? D'où ?","2. Vrai/Faux : Yasmine est en vacances chez sa tante. Justifie.","3. Cite 3 choses qu'on trouve au souk.","4. Que se passe-t-il le soir sur la place Jemaa el-Fna ?","5. Qu'est-ce que la harira ?","6. Combien de temps durera l'excursion dans le désert ?","7. ✍️ Production : réponds à Yasmine (5 lignes) en racontant un voyage."],1.5,21)
b=btn(s,9.8,6.1,2.4,0.7,"Correction",TERRE,16); links.append((b,"cec"))
s=slide("C.E. — Correction","",key="cec")
bullets(s,["1. Yasmine écrit à Léa, depuis Marrakech (Maroc).","2. FAUX : elle est chez sa grand-mère.","3. Épices, tapis, lanternes en cuivre.","4. Conteurs, musiciens gnaoua, petits restaurants.","5. Une soupe (traditionnelle).","6. Trois jours.","7. Production libre — critères : formule d'appel, passé composé, 3 détails, formule finale."],1.5,21)

# culture
s=slide("7. Cuisine & culture : devine l'image !","15'",key="cult",notes="Clique sur chaque carte pour révéler la réponse.")
cards=[("🍲","Je cuis dans un plat en terre au couvercle pointu.","tajine"),("🥣","Graine de semoule, je suis servi le vendredi.","couscous"),("🥾","Chaussure en cuir, souvent jaune.","babouche"),("🏺","Marché traditionnel couvert.","souk")]
for i,(e,d,k) in enumerate(cards):
    x=0.6+i*3.15; rect(s,x,1.6,2.95,3.6,BLANC,MSO_SHAPE.ROUNDED_RECTANGLE)
    txt(s,x,1.7,2.95,1,e,48,NOIR,False,PP_ALIGN.CENTER); txt(s,x+0.1,2.9,2.75,2,"« "+d+" »",16,NOIR,False,PP_ALIGN.CENTER)
    b=btn(s,x+0.4,5.4,2.15,0.7,"Qui suis-je ?",[TERRE,VERT,BLEU,OR][i],16); links.append((b,"cultR"))
s=slide("Réponses : tajine • couscous • babouche • souk","",key="cultR")
bullets(s,["🍲 Le TAJINE — plat et cuisson à l'étouffée","🥣 Le COUSCOUS — classé au patrimoine de l'UNESCO (2020)","🥾 La BABOUCHE","🏺 Le SOUK","🎵 Bonus : raï, chaâbi, gnaoua — quelles musiques connais-tu ?"])

# jeux
s=slide("8. Jeux de mots : le mot mystère","15'",key="jeux",notes="Les équipes devinent lettre par lettre (pendu).")
txt(s,0.6,1.7,12,1,"_ _ _ _ _ _ _   (7 lettres)",48,BLEU,True,PP_ALIGN.CENTER)
txt(s,0.6,3.0,12,1,"Indice : arbre des oasis qui donne des dattes 🌴",24,NOIR,False,PP_ALIGN.CENTER)
b=btn(s,5.2,4.4,2.9,0.9,"Révéler",TERRE,22); links.append((b,"jeuxR"))
s=slide("Réponse du mot mystère","",key="jeuxR")
txt(s,0.6,2.0,12,1.5,["D A T T I E R  ✅","Le palmier-dattier : l'arbre roi des oasis 🌴"],40,TERRE,True,PP_ALIGN.CENTER)
s=slide("Jeu : associe le mot et la définition","")
pairs=[("Médina","Vieille ville entourée de remparts"),("Casbah","Citadelle, quartier fortifié"),("Oasis","Point d'eau dans le désert"),("Hammam","Bain de vapeur traditionnel"),("Riad","Maison avec jardin intérieur")]
import random; random.seed(3); defs=[d for _,d in pairs]; random.shuffle(defs)
for i,(m,_) in enumerate(pairs):
    btn(s,0.8,1.6+i*1.0,3,0.8,m,BLEU,20); r=rect(s,6.0,1.6+i*1.0,6.6,0.8,BLANC); txt(s,6.2,1.7+i*1.0,6.4,0.6,"ABCDE"[i]+". "+defs[i],18)
s.notes_slide.notes_text_frame.text="Solution : "+" ; ".join(f"{m} → {d}" for m,d in pairs)

# défi
s=slide("9. Défi final : 60 secondes chrono !","10'",key="defi",notes="Chaque équipe cite le plus de mots liés au Maghreb en 60 s. 1 mot = 1 point.")
txt(s,0.6,1.7,12,1,"⏳ 60 s",80,ROUGE,True,PP_ALIGN.CENTER)
bullets(s,["Cite le plus de mots possibles sur le Maghreb (villes, plats, paysages…)","1 mot correct = 1 point • pas de répétition !","🏆 L'équipe gagnante choisit la musique de fin"],3.6,22)

s=slide("10. Bilan : qu'as-tu appris ?","10'",key="bilan")
bullets(s,["🎯 3 choses que j'ai apprises…","❤️ 2 choses que j'ai aimées…","❓ 1 question qu'il me reste…","✍️ Ticket de sortie : écris une carte postale du Maghreb (3 phrases)"],1.6,24)
txt(s,0.6,5.6,12,1,"Choukran • Tanemmirt • Merci ! 🌅",36,TERRE,True,PP_ALIGN.CENTER)

for sh,k in links: sh.click_action.target_slide=keys[k]
P.save("/home/user/Paquereauman/maghreb/maghreb-interactif.pptx"); print(len(P.slides))
