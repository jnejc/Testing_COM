# NMR string manipulation and list generation methods



def ToN(string):
    '''Converts string number to a float'''
    suffixes = {
        'n':0.000000001,
        'u':0.000001,
        'm':0.001,
        's':1
    }
    check = ('n','u','m','s')
    if string.endswith(check):
        #print(string)
        return float(string[:-1])*suffixes[string[-1]]
    else: return float(string)



def ToStr(num):
    '''Converts number to string format with suffix'''
    suffixes = {
        0:'s',
        1:'m',
        2:'u',
        3:'n',
        4:'p'
    }

    # check for zero, return without suffix
    if num == 0:
        return '{:.3g}'.format(num) + suffixes[0]

    # find correct suffix
    count = 0
    while num < 1 and count <4:
        count += 1
        num = num*1000
    return '{:.3g}'.format(num) + suffixes[count]



def Geometric_list(start, stop, n, shuffle=True):
    '''Generates the string of d5 values to measure T1'''
    print(start, stop)
    # Converts if input data are strings
    if type(start) == str:
        start = ToN(start)
    if type(stop) == str:
        stop = ToN(stop)
    if type(n) == str:
        n = int(n)

    # Creates geometric series
    mult = (stop/start)**(1./(n-1))
    print(mult)

    #table = [ToStr(start*(mult**i)) for i in range(n)]
    table = [ToStr((start*(mult**i))*10**8//10/10**7) for i in range(n)]
    print(table)


    # Parcels into 4 segments
    if shuffle:
        temp = []
        for i in [0,2,1,3]:
            for j in range(n//4 + int(n%4>i)):
                temp.append(4*j + i)
        temp = temp[:n]
        print(temp)
        table = [table[i] for i in temp]

    return ' '.join(table)



def Arithmetic_list(step, n, start=0, shuffle=False):
    '''Generates the string of d5 values to measure T1'''
    print(step)
    # Converts if input data are strings
    if type(step) == str:
        step = ToN(step)
    if type(n) == str:
        n = int(n)
    if type(start) == str:
        start = ToN(start)

    # Creates arithmetic series
    table = [ToStr((i*step + start)*10**8//10/10**7) for i in range(n)]

    # Parcels into 4 segments
    if shuffle:
        temp = []
        for i in [0,2,1,3]:
            for j in range(n//4 + int(n%4>i)):
                temp.append(4*j + i)
        temp = temp[:n]
        print(temp)
        table = [table[i] for i in temp]

    return ' '.join(table)



if __name__ == "__main__":

    print(ToStr(0))
    print(ToStr(0.0000000000000000000000000000001))
    print(ToStr(0.000000001))

    print(ToN('13u'))
    print(ToN('13.52m'))

    print(Geometric_list('20u', '20s', 21, True))
    print(Arithmetic_list('10u', 20, True))


