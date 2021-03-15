try:
    print("hello world")
    a=10
    b=0
    print("the sum is:%d", a/b)

except ZeroDivisionError:
    print("cannot perform zero division")

finally:
    print("operation completed")
