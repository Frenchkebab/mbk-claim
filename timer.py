import time
from random import random, randint

def sleep_timer_second(min, max):
    range = max - min
    time.sleep(min + random()*range)

def sleep_timer_minute(min, max):
    min = min * 60
    max = max * 60
    range = max - min
    time.sleep(min + random()*range)

def waitLoading():
    time.sleep(2)