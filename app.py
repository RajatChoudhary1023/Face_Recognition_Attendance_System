from flask import Flask,jsonify
import threading
from flask_cors import CORS
import cv2
import numpy as np
from PIL import Image
import os
import gspread
from gspread_formatting import CellFormat, format_cell_range, Color
from dotenv import load_dotenv

app=Flask(__name__)
CORS(app)
@app.route('/',methods=['GET'])

def show():


# Authenticate and open the spreadsheet
    gc = gspread.service_account(filename=os.getenv('SERVICE_ACCOUNT_PATH'))
    worksheet = gc.open_by_key(os.getenv('GOOGLE_SHEET'))

    current_sheet = worksheet.sheet1
    cell_values = current_sheet.get_all_values()
    cell_format = CellFormat(backgroundColor=Color(0.8, 1.0, 0.8))
    path='data'                          

    recognizer = cv2.face.LBPHFaceRecognizer_create()
    detector = cv2.CascadeClassifier("haarcascade_frontalface_default.xml")


    def imgsandlables (path):
        imagePaths = [os.path.join(path,i) for i in os.listdir(path)]     
        indfaces=[]
        ids = []
        for imagePath in imagePaths:
            img = Image.open(imagePath).convert('L') # grayscale
            imgnp = np.array(img,'uint8')
            id = int(os.path.split(imagePath)[-1].split(".")[0])
            
            faces = detector.detectMultiScale(imgnp)
            for (x,y,w,h) in faces:
                indfaces.append(imgnp[y:y+h,x:x+w])
                ids.append(id)
        return indfaces,ids



    faces,ids = imgsandlables (path)
    recognizer.train(faces, np.array(ids))
    checkface=[]
    id = 0
    names = [] 
    present=[]

    cam= cv2.VideoCapture(0)

    while True:
        _, img =cam.read()
        img = cv2.flip(img, 1) 
        gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
        faces = detector.detectMultiScale( gray, scaleFactor = 1.3, minNeighbors = 5,)

        index=0
        for(x,y,w,h) in faces:
            cv2.rectangle(img, (x,y), (x+w,y+h), (0,255,0), 2)

            id, confidence = recognizer.predict(gray[y:y+h,x:x+w])
            # Check if confidence is less them 100 ==> "0" is perfect match 
            if (confidence < 100):
                percent=round(100-confidence)
                index=id
                id = names[id]
                confidence = "  {0}%".format(round(100 - confidence))
                if percent>60:
                    print("Face Recognized Successfully("+str(id)+")")
                    checkface[index]+=1
            # else:
            #     id = "I dont know who this is :)"
            #     confidence = "  {0}%".format(round(100 - confidence))
            
            cv2.putText(img, str(id), (x-5,y-5), cv2.FONT_HERSHEY_SIMPLEX, 1, (255,255,0), 2)
            cv2.putText(img, str(confidence), (x+5,y+h-5), cv2.FONT_HERSHEY_SIMPLEX, 1, (255,255,0), 2) 


        cv2.imshow('camera',img) 

        if checkface[index]>2:
            if id not in present:
                print("Hi "+str(id)+"!")
                for i in range(16):
                    for j in range(3):
                        row=cell_values[i]
                        name=row[j]
                        if name==id:
                            current_sheet.update_cell(i+1,3,"Present")
                            format_cell_range(current_sheet, f"C{i+1}", cell_format)
                            print("Marked Present")
                            present.append(id)
        if cv2.waitKey(10) & 0xFF==ord('q'):
            break
    cam.release()
    cv2.destroyAllWindows()
            # row = cell_values[5]
            # roll_no = row[0]
            # roll_number = int(roll_no[2:])  # Assumes roll numbers are in format "TC" followed by digits
    return None
def start():
    threading.Thread(target=show).start()
    return jsonify(status="Camera Started")
app.run(debug=True)

