# looking for specified location to create a point and fill the attribute from excel (photo name)
# https://gis.stackexchange.com/questions/9433/extracting-coordinates-of-polygon-vertices-in-arcmap
# Author - Kyaw Htet Aung & Kyaw Naing Win(OneMap Myanmar)
# Date - 20200813

#Feature must be WGS84 Datum, otherwise calculated value will be wrong

import pandas as pd
import arcpy
import os

fcOrig = arcpy.GetParameterAsText(0)
colName = arcpy.GetParameterAsText(1)
tn = arcpy.GetParameterAsText(2)
col1 = arcpy.GetParameterAsText(3)
col2 = arcpy.GetParameterAsText(4)
col3 = arcpy.GetParameterAsText(5)
path = arcpy.GetParameterAsText(6)
name = arcpy.GetParameterAsText(7)
isflip = arcpy.GetParameterAsText(8)

arcpy.env.overwriteOutput = True 

'''
fc = r"D:\DOH\doh_lrs_2d3d_analysis\track_2D_10mile_utm47.shp"
colName = "rdId"
tn = r"D:\DOH\photo_location.xlsx\Sheet1$"
path = r"D:\DOH\Finding_Photo_Location"
name = "test010"
'''

doneList = [] #create a list not to add duplicate

#change the Spatial Reference to projected System (UTM 47 N)
arcpy.CreateFileGDB_management("C:/", "fGDB.gdb")
fc = r"C:\fGDB.gdb\Temp_fc"
outCS = arcpy.SpatialReference("WGS 1984 UTM Zone 47N")
arcpy.Project_management(fcOrig,fc,outCS)

if name.endswith(".shp"):
    name = name
else:
    name = name + ".shp"
file_name = path + "\\" + name

#creating feature class
arcpy.CreateFeatureclass_management(path, name, "POINT", has_z="ENABLED", spatial_reference=fc)

#adding new columns
arcpy.AddField_management(file_name,"X","DOUBLE")
arcpy.AddField_management(file_name,"Y","DOUBLE")
arcpy.AddField_management(file_name,"threeD_len","DOUBLE")
arcpy.AddField_management(file_name,"twoD_len","DOUBLE")
arcpy.AddField_management(file_name,"phName","TEXT")
arcpy.AddField_management(file_name,"rdId","TEXT")

#what attributes will be added when every new feature is added
cursor0 = arcpy.da.InsertCursor(file_name, ["SHAPE@XY","X","Y","twoD_len","phName","rdId"])

#reading the excel file by panda
#df = pd.read_excel(r"D:\DOH\photo_location.xlsx")

file, sheet = os.path.split(tn)
df = pd.read_excel(file, sheet_name= sheet)

cols = list(df.columns)
id0 = cols.index(col1)
id1 = cols.index(col2)
id2 = cols.index(col3)
#id0 = cols.index("road_id")
#id1 = cols.index("len")
#id2 = cols.index("photo")

for index0, row1 in df.iterrows():
    row1 = list(row1)
    rdId = row1[int(id0)]
    rdLen = row1[int(id1)]
    phName = row1[int(id2)]

    expression = "\"{}\" = \'{}\'".format(colName,rdId)  #to use in the where_clause(filtering)
    sum2D = 0

    #reading the feature class
    with arcpy.da.SearchCursor(fc, ['SHAPE@'],where_clause= expression) as cursor:
        for row in cursor:
            array1 = row[0].getPart()
            if isflip == "true":
                start = row[0].pointCount -1
                stop = -1
                step = -1

            else:
                start = 0
                stop = row[0].pointCount
                step = 1
            #for index in range(row[1].pointCount,-1,-1):
            for index in range(start,stop,step):
                if index:
                    pnt = array1.getObject(0).getObject(index) #get 2nd point
                    pnt0 = array1.getObject(0).getObject(index - 1) #get 1st point
                    x1 = pnt.X
                    y1 = pnt.Y
                    x0 = pnt0.X
                    y0 = pnt0.Y
                    dx = x1 - x0
                    dy = y1 - y0
                    dist2D = math.sqrt((dx * dx) + (dy * dy))
                    #if not(rdId in doneList): #blocking the duplicate values
                        #if sum2D != rdLen :
                            #cursor0.insertRow([(x0, y0), x0, y0, (sum2D),"",rdId]) #add old points

                    if sum2D < rdLen < (sum2D + dist2D): #specifying the location (the extect point as wanted)
                        s = rdLen -sum2D
                        dX = x1 - x0
                        dx = dX * (s / dist2D)
                        dY = y1 - y0
                        dy = dY * (s / dist2D)
                        X = x0 + dx
                        Y = y0 + dy
                        cursor0.insertRow([(X, Y), X, Y, rdLen,phName,rdId]) #add new points and some attributes
                    sum2D = sum2D + dist2D
            #doneList.append(rdId)

print("code was successfully run")

#add the feature to current ArcMap#add the feature to current ArcMap
arcpy.AddMessage("Finished")
mxd = arcpy.mapping.MapDocument("CURRENT")
layer = arcpy.mapping.Layer(file_name)
df1 = arcpy.mapping.ListDataFrames(mxd)[0]
arcpy.mapping.AddLayer(df1,layer)