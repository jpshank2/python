__counter = 0

def suml(theList):
    global __counter
    __counter += 1
    theSum = 0
    for element in theList:
        theSum += element
    return theSum

def prodl(theList):
    global __counter
    __counter += 1
    prod = 1
    for element in theList:
        prod *= element
    return prod

if __name__ == "__main__":
    print("I prefer to be a module, but I can do some tests for you.")
    my_list = [i+1 for i in range(5)]
    print(suml(my_list) == 15)
    print(prodl(my_list) == 120)