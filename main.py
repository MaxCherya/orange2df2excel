from orange2df2excel import gen_encryption_key, rederive_key

# x = gen_encryption_key("banana")
# print(x)

y = rederive_key('banana', salt=b'\x0es\xec\x0cC\x8b\xd5<\x99\x9c\xd7\xf2C$\xae?8\xf3R\xee\x16\xbb\x04N\x1c\x1b\xdd>\xe65\x84X')
print(y)