with open('mk.txt', 'r', encoding='windows-1254') as f:
    line = f.readline().rstrip('\n')
    print(f"Line length: {len(line)}")
    print("   " + "".join([str(i % 10) for i in range(len(line))]))
    print("   " + line)
    print("TDE-Sos1 (40):", repr(line[54:94]))
    print("Sosyal-2 (46):", repr(line[94:140]))
    print("Matematik(40):", repr(line[140:180]))
    print("Fen      (40):", repr(line[180:220]))
