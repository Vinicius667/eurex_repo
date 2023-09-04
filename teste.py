def a():
    global t 
    t = '1'
    b()
    print(v)

def b():
    global v
    v = 1
a()
print(t)
print(v)