import chardet

# Öppna fil och gissa encoding
with open('sie_to_pnl/data/IB.csv', 'rb') as f:
    result = chardet.detect(f.read())

print(result)
