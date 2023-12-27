#vorname_list = ['Jon', 'Daniel', 'Alex', 'Britta', 'Tim', 'Andre', 'Sven', 'Daniel']
#sort_vorname_list = sorted(vorname_list)
#vorname_list_finish = list(set(sort_vorname_list))
#
#print(sort_vorname_list)
#print(vorname_list_finish)
import time

my_list = ['Jon', 'Daniel', 'Alex', 'Britta', 'Tim', 'Andre', 'Sven', 'Daniel']
sorted_list = sorted(set(my_list))

print(sorted_list) # выводим уникальные элементы в отсортированном порядке

import threading

def process1():
    for i in range(1,20,2):
        print(i)
        time.sleep(1)

def process2():
    for i in range(0,20,2):
        print(i)
        time.sleep(1)

thread1 = threading.Thread(target=process1)
thread2 = threading.Thread(target=process2)

thread1.start()
thread2.start()