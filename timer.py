import time
from random import random, randint

def sleep_timer_second(min, max):
    range = max - min
    time.sleep(min + random()*range)

def waitLoading():
    time.sleep(2)