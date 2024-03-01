import datetime
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
import openpyxl
from pylsl import StreamInlet, resolve_stream

data = []  # List to store data points
mode = []  # List to store modes (not used in the provided code)
count = 0  # Counter variable
timeData = 0
# Initialize camera
cam = cv2.VideoCapture(0, cv2.CAP_DSHOW)


# Class for creating blinking lights using Turtle graphics
class BlinkingLights:
    def __init__(self, data_list):
        # Initialize Turtle graphics window
        self.screen = turtle.Screen()
        self.screen.title('Blinking Lights')
        self.screenTk = self.screen.getcanvas().winfo_toplevel()
        self.screenTk.attributes("-fullscreen", True)
        self.screen.bgcolor("black")
        self.round = 0  # Counter for rounds of blinking
        self.brightness = 0.0  # Initial brightness of lights
        self.hues = [0.0, 0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6]  # List of hues for colors
        self.dots = []  # List to store Turtle objects representing lights
        self.blinking = []  # List to store blinking status of each light
        self.data_list = data_list  # Reference to the shared data list
        self.workbook = xlsxwriter.Workbook("AllAboutdata.xlsx")  # Workbook for storing data
        self.worksheet = self.workbook.add_worksheet("firstSheet")  # Worksheet for storing data
        # Write column headers in the worksheet
        self.worksheet.write(0, 0, "Top Left (x,y,time stamp)")
        self.worksheet.write(0, 1, "Top Center (x,y,time stamp)")
        self.worksheet.write(0, 2, "Top Right (x,y,time stamp)")
        self.worksheet.write(0, 3, "Center Left (x,y,time stamp)")
        self.worksheet.write(0, 4, "Center (x,y,time stamp)")
        self.worksheet.write(0, 5, "Center Right (x,y,time stamp)")
        self.worksheet.write(0, 6, "Bottom Left (x,y,time stamp)")
        self.worksheet.write(0, 7, "Bottom Center (x,y,time stamp)")
        self.worksheet.write(0, 8, "Bottom Right (x,y,time stamp)")
        self.row = 0  # Row index for writing data
        self.col = 0  # Column index for writing data

    @staticmethod
    def create_dot(position):
        # Create a Turtle object representing a light dot
        dot = turtle.Turtle()
        dot.shape("circle")
        dot.shapesize(1, 1)
        dot.goto(position)
        dot.penup()
        return dot

    def change_brightness_sequence(self, index):
        # Method to change the brightness of lights in sequence
        self.round += 1
        global count
        count = self.round
        if self.round > 9:
            # If completed 9 rounds, close the Turtle graphics window and workbook
            time.sleep(3)
            self.screen.bye()
            self.workbook.close()

        # Toggle brightness between 0.1 and 0.9
        self.brightness = 0.9 if self.brightness < 1.0 else 0.1
        rgb = colorsys.hsv_to_rgb(self.hues[index], 1, self.brightness)
        self.dots[index].color(rgb)

        # Write data to the worksheet for each light
        column = 1
        for line in self.data_list:
            self.worksheet.write(column, self.round - 1, line)
            column += 1

        # Clear the shared data list
        data_list.clear()

        # Schedule the next change in brightness
        if self.blinking[index]:
            self.screen.ontimer(functools.partial(self.change_brightness_sequence, index), 100)
            self.data_list.append(self.dots[index].position())
        else:
            self.screen.ontimer(functools.partial(self.hide_dot, index), 5000)

    def hide_dot(self, index):
        # Method to hide a light dot and move to the next dot
        self.dots[index].hideturtle()
        next_index = (index + 1) % len(self.dots)
        self.dots[next_index] = self.create_dot(self.dots[next_index].position())
        self.screen.ontimer(functools.partial(self.change_brightness_sequence, next_index), 200)

    def start_blinking_lights(self):
        # Method to start the blinking lights sequence
        # Create light dots at specific positions
        self.dots = [self.create_dot((-800 + (i % 3) * 800, 500 - int(i / 3) * 500)) for i in range(9)]
        self.blinking = [False] * len(self.dots)
        # Start the sequence by changing brightness of the first light
        self.screen.ontimer(functools.partial(self.change_brightness_sequence, 0), 3000)
        self.screen.mainloop()


# Function to start the Turtle graphics for blinking lights
def start_turtle_graphics(data_list):
    blinking_lights = BlinkingLights(data_list)
    blinking_lights.start_blinking_lights()


# Function to start webcam interaction for detecting facial landmarks
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
                    current_time = datetime.datetime.now()
                    global timeData
                    timeData = current_time
                    if data_list is not None:
                        data_list.append('X: ' + str(screen_x) + ' Y: ' + str(screen_y) + " Time:  " + str(current_time))
                    if ( count > 9):
                        worksheet.write(col, row,
                                        'X: ' + str(screen_x) + ' Y: ' + str(screen_y) + " Time: " + str(current_time))
                        col += 1
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


def lsl_streaming():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    header=["Timestamp","Data"]
    sheet.append(header)
    workbook.save("realtime_data.xlsx")
    # first resolve an EEG stream on the lab network
    print("looking for an EEG stream...")
    streams = resolve_stream('type', 'Event')

    # create a new inlet to read from the stream
    inlet = StreamInlet(streams[0])

    while True:
        # get a new sample (you can also omit the timestamp part if you're not
        # interested in it)
        sample, timestamp = inlet.pull_sample()
        print(timestamp, sample)

        # Write data to the worksheet
        for col, data_point in enumerate(sample):
            sheet.append([data_point+" TimeStamp: "+str(timeData)])
        workbook.save("realtime_data.xlsx")
        # Sleep for a short time to avoid excessive CPU usage
        time.sleep(0.1)

    # Close the workbook when done
    workbook.close()


if __name__ == "__main__":
    data_list = []  # Shared list for storing data points
    # Create threads for running Turtle graphics, webcam interaction, and LSL streaming concurrently
    turtle_thread = threading.Thread(target=start_turtle_graphics, args=(data_list,))
    webcam_thread = threading.Thread(target=start_webcam_interaction, args=(data_list,))
    lsl_thread = threading.Thread(target=lsl_streaming, args=())
    # Start the threads
    turtle_thread.start()
    webcam_thread.start()
    lsl_thread.start()
    # Wait for threads to finish
    turtle_thread.join()
    webcam_thread.join()
    lsl_thread.join()