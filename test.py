cell = 'A21'


n = 0
char = []
char[n] = list(cell)
n = 1
for char in cell :
    char[n] = list(cell)[n]
    n += 1
    print(char[char])

