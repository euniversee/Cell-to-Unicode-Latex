# Cell-to-Unicode-Latex
Mengubah dari cell di excel menjadi rumus latex / unicode, dari rumus tersebut bisa langsung diubah menjadi equation di word (biar gampang ngelapraknya ga ngetik equation satu-satu)

How to Use:
1. Buka Excel, tekan Alt + F11 (gunakan Fn jika tidak bisa membuka VBA)
2. Tab File, import File
3. Kembali ke excel, gunakan =FormulaToUnicode(cell yang dituju) atau =FormulaToLatex(cell)
4. Akan tergenerate rumus latex/unicode
5. copy Pergi ke Word, Paste, Select, Pencet Insert Equation
6. jika latex tekan latex pada menu equation
7. klik convert Current - Professional
8. ヽ(´▽`)/

note:
penggunaan masih terbatas pada equation sederhana, untuk unicode sudah ada variabel awal, misal l^2=rumus, l^2 di cell a20 dan rumus di cell b20, maka cell a20 akan ikut ke equation
