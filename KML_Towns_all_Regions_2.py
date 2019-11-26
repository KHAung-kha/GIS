#required modules : shutil, pandas, time

#turn the crrodinates from excel to kml with village names, state by state
#define the head, tail, style, point start and point end of kml first
#merge them with coordinate from excel
import shutil
import pandas as pd
import time as tm
st_time = tm.time()
head = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><kml xmlns=\"http://www.opengis.net/kml/2.2\" xmlns:gx=\"http://www.google.com/kml/ext/2.2\" xmlns:kml=\"http://www.opengis.net/kml/2.2\" xmlns:atom=\"http://www.w3.org/2005/Atom\"><Document>"
style = "<StyleMap id=\"msn_placemark_circle\"><Pair><key>normal</key><styleUrl>#sn_placemark_circle</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#sh_placemark_circle_highlight</styleUrl></Pair></StyleMap><Style id=\"sh_placemark_circle_highlight\"><IconStyle><scale>1.2</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href></Icon></IconStyle><ListStyle></ListStyle></Style><Style id=\"sn_placemark_circle\"><IconStyle><scale>1.2</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon></IconStyle><ListStyle></ListStyle></Style><Folder>"
tail = "</Folder></Document></kml>"
ptend = "</coordinates></Point></Placemark>"
df = pd.read_excel(r"C:\Users\Thinkpad\Downloads\Compressed\Myanmar_PCodes_Release-IX_Sep2019_Countrywide\Myanmar PCodes Release-IX_Sep2019_Countrywide.xlsx",sheet_name="_07_Villages")
df = df[['SR_Name_Eng','Village_Name_Eng','Latitude','Longitude']] #filtering the wanted colums only
df= df.dropna() #removing NA rows
#print(df.shape)
#print(df.head())
alt = 0
stName = df['SR_Name_Eng']
stName= set(stName)
print(stName)
for eachst in stName:
    df1=df.loc[(df['SR_Name_Eng'] == eachst)]
    #print(df1.shape)
    points = ""
    for index,row in df1.iterrows():
        name = row['Village_Name_Eng']
        ptstart = f"<Placemark><name>{name}</name><styleUrl>#msn_placemark_circle</styleUrl><Point><gx:drawOrder>1</gx:drawOrder><coordinates>"
        lat = row['Latitude']
        lon = row['Longitude']
        coordinate = str(lon) + "," + str(lat) + "," + str(alt) + " "
        points = points + ptstart+ coordinate + ptend
    placemark = head + style + points + tail
    Namefile = "kml_creation_KHA_" + eachst + ".kml"
    kml_file = open(Namefile,"w",encoding="utf-8")
    kml_file.write(placemark)
    kml_file.close()
    print(eachst + " has been written!")
    source = "E:\Python_scrpits\python_files\\" + Namefile
    distination = "E:\KMLs\\" + Namefile
    tm.sleep(1)
    shutil.move(source, distination) #moving the files to other directory
end_time = tm.time()
print(str(round(end_time - st_time)))
print("finished")
import pyttsx3 as pt
sound = pt.init()
sound.say("Hi, your code is successfully run and finished")
sound.runAndWait()