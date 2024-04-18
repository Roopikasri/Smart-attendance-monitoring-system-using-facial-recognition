from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import csv
import time
from datetime import datetime
import os

from win32com.client import Dispatch

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
face_cascade=cv2.CascadeClassifier("D:\Dev\Face-rec\DATA\haarcascade_frontalface_default.xml")
with open('Face-rec/names.pkl', 'rb') as f:
    LABELS=pickle.load(f)
with open('Face-rec/faces_data.pkl', 'rb') as f:
    FACES=pickle.load(f)
knn=KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)
COL_NAMES = ['NAME', 'TIME']
while True:
    ret,frame=cap.read()
    gray=cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces=face_cascade.detectMultiScale(gray, 1.3 ,5)
    for (x,y,w,h) in faces:
        crop_img=frame[y:y+h, x:x+w, :]
        crop_img=frame[y:y+h, x:x+w, :]
        resized_img=cv2.resize(crop_img, (50,50)).flatten().reshape(1,-1)
        output=knn.predict(resized_img)
        ts=time.time()
        date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp=datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist=os.path.isfile("D:\\Dev\\Face-rec\\attendance\\attendance" + date + ".csv")
        cv2.putText(frame, str(output[0]), (x,y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
        cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 1)
        attendance=[str(output[0]), str(timestamp)]
    cv2.imshow("frame",frame)
    k=cv2.waitKey(1)
    if k==ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("D:\\Dev\\Face-rec\\attendance\\attendance" + date + ".csv", "+a") as csvfile:
                writer=csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("D:\\Dev\\Face-rec\\attendance\\attendance" + date + ".csv", "+a") as csvfile:
                writer=csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()
    if k==ord('q'):
        break
cap.release()
cv2.destroyAllWindows()