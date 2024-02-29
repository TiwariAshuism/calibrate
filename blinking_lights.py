import time
import turtle
import colorsys
import functools
import cv2
import mediapipe as mp
import pyautogui

class BlinkingLights:
    def __init__(self):
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
        if self.round > 9:
            time.sleep(3)
            self.screen.bye()

        self.brightness = 0.9 if self.brightness < 1.0 else 0.1
        rgb = colorsys.hsv_to_rgb(self.hues[index], 1, self.brightness)
        self.dots[index].color(rgb)

        if self.blinking[index]:
            self.screen.ontimer(functools.partial(self.change_brightness_sequence, index), 100)
        else:
            self.screen.ontimer(functools.partial(self.hide_dot, index), 1000)

    def hide_dot(self, index):
        self.dots[index].hideturtle()
        next_index = (index + 1) % len(self.dots)
        self.dots[next_index] = self.create_dot(self.dots[next_index].position())
        self.screen.ontimer(functools.partial(self.change_brightness_sequence, next_index), 200)

    def start_blinking_lights(self):
        self.dots = [self.create_dot((-400 + (i % 3) * 400, 300 - int(i / 3) * 300)) for i in range(9)]
        self.blinking = [False] * len(self.dots)
        self.screen.ontimer(functools.partial(self.change_brightness_sequence, 0), 50)
        self.screen.mainloop()

def start_turtle_graphics():
    blinking_lights = BlinkingLights()
    blinking_lights.start_blinking_lights()

def start_webcam_interaction():
    cam = cv2.VideoCapture(0)
    face_mesh = mp.solutions.face_mesh.FaceMesh(refine_landmarks=True)
    screen_w, screen_h = pyautogui.size()
    while True:
        _, frame = cam.read()
        frame = cv2.flip(frame, 1)
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        output = face_mesh.process(rgb_frame)
        landmark_points = output.multi_face_landmarks
        frame_h, frame_w, _ = frame.shape
        if landmark_points:
            landmarks = landmark_points[0].landmark
            for id, landmark in enumerate(landmarks[474:478]):
                x = int(landmark.x * frame_w)
                y = int(landmark.y * frame_h)
                cv2.circle(frame, (x, y), 3, (0, 255, 0))
                if id == 1:
                    screen_x = int(screen_w * landmark.x)
                    screen_y = int(screen_h * landmark.y)
                    print("Cursor Location:", screen_x, screen_y)  # Print cursor location
            left = [landmarks[145], landmarks[159]]
            for landmark in left:
                x = int(landmark.x * frame_w)
                y = int(landmark.y * frame_h)
                cv2.circle(frame, (x, y), 3, (0, 255, 255))
            if (left[0].y - left[1].y) < 0.004:
                pyautogui.click()
                pyautogui.sleep(1)
        cv2.imshow('Eye Controlled ', frame)
        cv2.waitKey(1)
