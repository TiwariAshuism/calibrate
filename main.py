import threading
import time
import turtle
import colorsys
import functools
import cv2
import mediapipe as mp
import pyautogui
import xlsxwriter
import statistics
data = []
mode = []
count=0
cam = cv2.VideoCapture(0,cv2.CAP_DSHOW)
class BlinkingLights:
    def __init__(self, data_list):
        self.screen = turtle.Screen()
        self.screen.title('Blinking Lights')
        self.screenTk = self.screen.getcanvas().winfo_toplevel()
        self.screenTk.attributes("-fullscreen", True)
        self.screen.bgcolor("black")
        self.round = 0
        self.brightness = 0.0
        self.hues = [0.0, 0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6]
        self.dots = []
        self.blinking = []
        self.data_list = data_list
        self.workbook = xlsxwriter.Workbook("AllAboutdata.xlsx")
        self.worksheet = self.workbook.add_worksheet("firstSheet")
        self.worksheet.write(0, 0, "Point 1")
        self.worksheet.write(0, 1, "Point 2")
        self.worksheet.write(0, 2, "Point 3")
        self.worksheet.write(0, 3, "Point 4")
        self.worksheet.write(0, 4, "Point 5")
        self.worksheet.write(0, 5, "Point 6")
        self.worksheet.write(0, 6, "Point 7")
        self.worksheet.write(0, 7, "Point 8")
        self.worksheet.write(0, 8, "Point 9")
        self.row = 0
        self.col = 0

    @staticmethod
    def create_dot(position):
        dot = turtle.Turtle()
        dot.shape("circle")
        dot.shapesize(1, 1)
        dot.goto(position)
        dot.penup()
        return dot

    def change_brightness_sequence(self, index):
        self.round += 1
        global count
        count=self.round
        if self.round > 9:
            time.sleep(3)
            self.screen.bye()
            self.workbook.close()

        self.brightness = 0.9 if self.brightness < 1.0 else 0.1
        rgb = colorsys.hsv_to_rgb(self.hues[index], 1, self.brightness)
        self.dots[index].color(rgb)
        column=1
        for line in self.data_list:
            self.worksheet.write(column, self.round-1, line)
            column += 1

        data_list.clear()
        if self.blinking[index]:
            self.screen.ontimer(functools.partial(self.change_brightness_sequence, index), 100)
            self.data_list.append(self.dots[index].position())

        else:
            self.screen.ontimer(functools.partial(self.hide_dot, index), 500)

    def hide_dot(self, index):
        self.dots[index].hideturtle()
        next_index = (index + 1) % len(self.dots)
        self.dots[next_index] = self.create_dot(self.dots[next_index].position())
        self.screen.ontimer(functools.partial(self.change_brightness_sequence, next_index), 200)

    def start_blinking_lights(self):
        self.dots = [self.create_dot((-800 + (i % 3) * 800, 500 - int(i / 3) * 500)) for i in range(9)]
        self.blinking = [False] * len(self.dots)
        self.screen.ontimer(functools.partial(self.change_brightness_sequence, 0), 50)
        self.screen.mainloop()


def start_turtle_graphics(data_list):
    blinking_lights = BlinkingLights(data_list)
    blinking_lights.start_blinking_lights()


def start_webcam_interaction(data_list=None):
    face_mesh = mp.solutions.face_mesh.FaceMesh(refine_landmarks=True)
    screen_w, screen_h = pyautogui.size()
    workbook = xlsxwriter.Workbook("Alldata.xlsx")
    worksheet = workbook.add_worksheet("firstSheet")
    worksheet.write(0, 0, "Point 1")
    row = 0
    col = 0
    while True:
        ret, frame = cam.read()
        if not ret:
            print("Error: Unable to access the camera.")
            break
        frame = cv2.flip(frame, 1)
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        output = face_mesh.process(rgb_frame)
        landmark_points = output.multi_face_landmarks
        frame_h, frame_w, ret = frame.shape
        if landmark_points:
            landmarks = landmark_points[0].landmark
            for id, landmark in enumerate(landmarks[474:478]):
                x = int(landmark.x * frame_w)
                y = int(landmark.y * frame_h)
                cv2.circle(frame, (x, y), 3, (0, 255, 0))
                if id == 1:
                    screen_x = int(screen_w * landmark.x)
                    screen_y = int(screen_h * landmark.y)
                    current_time= time.time()
                    minute = int(current_time//60)%60
                    seconds = int(current_time%60)
                    if data_list is not None:
                        data_list.append('x: '+str(screen_x)+' y: '+str(screen_y)+"  "+str(minute)+":"+str(seconds) )
                    if(count>9):
                        worksheet.write(col, row, 'x: '+str(screen_x)+' y: '+str(screen_y)+"  "+str(minute)+":"+str(seconds))
                        col+=1
            left = [landmarks[145], landmarks[159]]
            for landmark in left:
                x = int(landmark.x * frame_w)
                y = int(landmark.y * frame_h)
                cv2.circle(frame, (x, y), 3, (0, 255, 255))
            if (left[0].y - left[1].y) < 0.004:
                pyautogui.click()
                pyautogui.sleep(1)

        cv2.imshow('Eye Controlled ', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            workbook.close()
            break


if __name__ == "__main__":
    data_list = []
    turtle_thread = threading.Thread(target=start_turtle_graphics, args=(data_list,))
    webcam_thread = threading.Thread(target=start_webcam_interaction, args=(data_list,))
    webcam_thread1=threading.Thread(target=start_webcam_interaction)

    turtle_thread.start()
    webcam_thread.start()

    turtle_thread.join()
    webcam_thread.join()