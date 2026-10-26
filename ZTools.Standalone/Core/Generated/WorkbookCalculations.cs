// Ported from migration/legacy; see tools/port_vba.py and migration/README.md.
using System;
namespace ZTools.Engine {
    internal sealed partial class SubmissionEngine {
        public void TransferData() {
            dynamic @abcd = "";
            dynamic @cell = null;
            dynamic @counter = 0;
            dynamic @counter1 = 0;
            dynamic @counter2 = 0;
            dynamic @countercheck = 0;
            dynamic @counterprogress = 0;
            dynamic @countersum = 0;
            dynamic @databaseoffset = 0;
            dynamic @gadatabasesheet = null;
            dynamic @i = 0;
            dynamic @j = 0;
            dynamic @k = 0;
            dynamic @lk = "";
            dynamic @name = "";
            dynamic @newpumptype = "";
            dynamic @nordia = "";
            dynamic @origpumptype = null;
            dynamic @other_weight = 0;
            dynamic @pctdone = 0;
            dynamic @progress = 0;
            dynamic @pump_type = null;
            dynamic @pump_type2 = null;
            dynamic @pump_type_all = null;
            dynamic @str1 = "";
            dynamic @str2 = "";
            dynamic @str2new = "";
            dynamic @str3 = "";
            dynamic @str4 = "";
            dynamic @str5 = "";
            dynamic @str6 = "";
            dynamic @x = 0;
            dynamic @xcheck = 0;
            dynamic @y = 0;
            dynamic @ycheck = 0;
            dynamic @z = 0;
            // Normalized rule line 1
            // Normalized rule line 2
            // Normalized rule line 3
            // Normalized rule line 4
            // Normalized rule line 5
            // Normalized rule line 6
            // Normalized rule line 7
            // Normalized rule line 8
            // Normalized rule line 9
            // Normalized rule line 10
            // Normalized rule line 11
            // Normalized rule line 17
            @i = ExcelApp.WorksheetFunction.CountA(Book.Worksheets["input"].Range["$n$21:$n$50"]);
            // Normalized rule line 19
            Book.Worksheets["Sheet1"].Range["$c$15"].Value2 = @i;
            // Normalized rule line 22
            for (@k = 0; Vb.Le(@k, Vb.Sub(@i, 1)); @k = Vb.Add(@k, 1)) {
                // Normalized rule line 23
                @origpumptype = "";
                // Normalized rule line 26
                if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range["$i$53"].Value2, "GB9113.1-2000"))) {
                    // Normalized rule line 27
                    for (@counter1 = 0; Vb.Le(@counter1, 2); @counter1 = Vb.Add(@counter1, 1)) {
                        // Normalized rule line 29
                        if (Vb.Bool(Vb.Eq(Vb.Instr(Book.Worksheets["Sheet1"].Range[Vb.Cat("C", Vb.Add(12, @counter1))].Value2, "P"), 0))) {
                            // Normalized rule line 32
                            throw new InvalidOperationException("计算流程提前终止，请检查输入参数及数据库匹配情况。");
                            // Normalized rule line 33
                        } else {
                            // Normalized rule line 34
                        }
                        // Normalized rule line 36
                    }
                    // Normalized rule line 37
                } else {
                    // Normalized rule line 39
                }
                // Normalized rule line 41
                if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2, ""))) {
                    // Normalized rule line 43
                    throw new InvalidOperationException("计算流程提前终止，请检查输入参数及数据库匹配情况。");
                    // Normalized rule line 44
                } else {
                    // Normalized rule line 46
                }
                // Normalized rule line 48
                if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("V", Vb.Add(21, @k))].Value2, ""))) {
                    // Normalized rule line 50
                    throw new InvalidOperationException("计算流程提前终止，请检查输入参数及数据库匹配情况。");
                    // Normalized rule line 51
                } else {
                    // Normalized rule line 53
                }
                // Normalized rule line 56
                if (Vb.Bool(Vb.Or(Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("V", Vb.Add(21, @k))].Value2, "6000V"), Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("V", Vb.Add(21, @k))].Value2, "10000V")))) {
                    // Normalized rule line 59
                } else {
                    // Normalized rule line 61
                }
                // Normalized rule line 63
                if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("AA", Vb.Add(21, @k))].Value2, ""))) {
                    // Normalized rule line 65
                    throw new InvalidOperationException("计算流程提前终止，请检查输入参数及数据库匹配情况。");
                    // Normalized rule line 66
                } else {
                    // Normalized rule line 68
                }
                // Normalized rule line 70
            }
            // Normalized rule line 73
            Book.Worksheets["Sheet1"].Activate();
            // Normalized rule line 81
            Book.Worksheets["Sheet1"].Range["$b$40:$i$69"].ClearContents();
            // Normalized rule line 82
            Book.Worksheets["Sheet1"].Range["$k$40:$cc$69"].ClearContents();
            // Normalized rule line 83
            Book.Worksheets["Sec-List"].Range["$f$10:$f$39"].ClearContents();
            // Normalized rule line 89
            @countersum = Vb.Mul(75, @i);
            // Normalized rule line 91
            @counterprogress = 1;
            // Normalized rule line 103
            for (@k = 0; Vb.Le(@k, Vb.Sub(@i, 1)); @k = Vb.Add(@k, 1)) {
                // Normalized rule line 104
                @origpumptype = "";
                // Normalized rule line 107
                Book.Worksheets["Sheet1"].Range[Vb.Cat("B", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("D", Vb.Add(21, @k))].Value2;
                // Normalized rule line 109
                Book.Worksheets["Sheet1"].Range[Vb.Cat("D", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("F", Vb.Add(21, @k))].Value2;
                // Normalized rule line 111
                Book.Worksheets["Sheet1"].Range[Vb.Cat("K", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("S", Vb.Add(21, @k))].Value2;
                // Normalized rule line 113
                Book.Worksheets["Sheet1"].Range[Vb.Cat("N", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("T", Vb.Add(21, @k))].Value2;
                // Normalized rule line 115
                Book.Worksheets["Sheet1"].Range[Vb.Cat("N", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("T", Vb.Add(21, @k))].Value2;
                // Normalized rule line 117
                @j = Vb.Len(Book.Worksheets["input"].Range[Vb.Cat("V", Vb.Add(21, @k))].Value2);
                // Normalized rule line 119
                Book.Worksheets["Sheet1"].Range[Vb.Cat("L", Vb.Add(40, @k))].Value2 = Vb.Mid(Book.Worksheets["input"].Range[Vb.Cat("V", Vb.Add(21, @k))].Value2, 1, Vb.Sub(@j, 1));
                // Normalized rule line 121
                Book.Worksheets["Sheet1"].Range[Vb.Cat("M", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("W", Vb.Add(21, @k))].Value2;
                // Normalized rule line 123
                Book.Worksheets["Sheet1"].Range[Vb.Cat("O", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("R", Vb.Add(21, @k))].Value2;
                // Normalized rule line 125
                Book.Worksheets["Sec-List"].Range[Vb.Cat("F", Vb.Add(10, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("O", Vb.Add(21, @k))].Value2;
                // Normalized rule line 129
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")), Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "NO")))) {
                    // Normalized rule line 133
                    if (Vb.Bool(Vb.Or(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "P"), 0), Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "p"), 0)))) {
                        // Normalized rule line 134
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2;
                        // Normalized rule line 135
                    } else {
                        // Normalized rule line 136
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2 = Vb.Cat("Plan ", Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2);
                        // Normalized rule line 137
                    }
                    // Normalized rule line 139
                } else {
                    // Normalized rule line 141
                    if (Vb.Bool(Vb.Or(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "P"), 0), Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "p"), 0)))) {
                        // Normalized rule line 142
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "+"), Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2);
                        // Normalized rule line 143
                    } else {
                        // Normalized rule line 144
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat(Vb.Cat("Plan ", Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2), "+"), Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2);
                        // Normalized rule line 145
                    }
                    // Normalized rule line 147
                }
                // Normalized rule line 154
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 156
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 159
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 167
                @pump_type_all = Book.Worksheets["input"].Range[Vb.Cat("N", Vb.Add(21, @k))].Value2;
                // Normalized rule line 170
                @j = Vb.Len(@pump_type_all);
                // Normalized rule line 171
                @str2 = "";
                // Normalized rule line 172
                for (@x = 1; Vb.Le(@x, @j); @x = Vb.Add(@x, 1)) {
                    // Normalized rule line 174
                    @str1 = Vb.Mid(@pump_type_all, @x, 1);
                    // Normalized rule line 175
                    if (Vb.Bool(Vb.Or(Vb.Eq(Vb.Asc(@str1), 32), Vb.Eq(Vb.Asc(@str1), 63)))) {
                        // Normalized rule line 176
                        @str1 = "";
                        // Normalized rule line 177
                    }
                    // Normalized rule line 179
                    if (Vb.Bool(Vb.Ne(@str1, " "))) {
                        // Normalized rule line 180
                        @str2 = Vb.Cat(@str2, @str1);
                        // Normalized rule line 181
                    } else {
                        // Normalized rule line 182
                        @str1 = "";
                        // Normalized rule line 183
                        @str2 = Vb.Cat(@str2, @str1);
                        // Normalized rule line 184
                    }
                    // Normalized rule line 185
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("C", Vb.Add(40, @k))].Value2 = @str2;
                    // Normalized rule line 187
                }
                // Normalized rule line 189
                if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2, ""))) {
                    // Normalized rule line 190
                    @str2new = Vb.Replace(@str2, Vb.Left(@str2, 2), Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2);
                    // Normalized rule line 191
                }
                // Normalized rule line 197
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 198
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 199
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 207
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(Vb.Mid(@str2, 3, 1), "O"), Vb.Eq(Vb.Mid(@str2, 3, 1), "o")), Vb.Eq(Vb.Mid(@str2, 3, 1), "G")), Vb.Eq(Vb.Mid(@str2, 3, 1), "g")))) {
                    // Normalized rule line 208
                    @str4 = Vb.Cat(Vb.Mid(@str2, 1, 2), Vb.Mid(@str2, 4, Vb.Sub(Vb.Len(@str2), 3)));
                    // Normalized rule line 209
                    @pump_type2 = Vb.Mid(@str2, 1, 3);
                    // Normalized rule line 210
                } else {
                    // Normalized rule line 211
                    @str4 = @str2;
                    // Normalized rule line 212
                    @pump_type2 = Vb.Mid(@str2, 1, 2);
                    // Normalized rule line 213
                }
                // Normalized rule line 214
                @pump_type = Vb.Mid(@str2, 1, 2);
                // Normalized rule line 215
                @nordia = Vb.Right(@str2, 3);
                // Normalized rule line 219
                if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")))) {
                    // Normalized rule line 220
                    @z = Vb.Len(@str4);
                    // Normalized rule line 221
                    for (@x = 1; Vb.Le(@x, @z); @x = Vb.Add(@x, 1)) {
                        // Normalized rule line 223
                        if (Vb.Bool(Vb.Eq(Vb.Mid(@str4, @x, 1), "-"))) {
                            break;
                        }
                        // Normalized rule line 225
                    }
                    // Normalized rule line 226
                    @str6 = Vb.Cat(Vb.Mid(@str4, 3, Vb.Sub(@x, 2)), Vb.Mid(@str4, Vb.Add(@x, 2), 3));
                    // Normalized rule line 228
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BR", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str6, Book.Worksheets["Data1"].Range["$c$10:$d$62"], 2, false);
                    // Normalized rule line 229
                } else {
                    // Normalized rule line 230
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BR", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 231
                }
                // Normalized rule line 234
                if (Vb.Bool(Vb.Eq(Vb.Left(@str4, 2), "CZ"))) {
                    // Normalized rule line 235
                    @lk = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets["CZ-BearingSize"].Range["$A$2:$B$41"], 2, false);
                    // Normalized rule line 236
                } else {
                    // Normalized rule line 237
                    @z = Vb.Len(@str4);
                    // Normalized rule line 238
                    for (@x = 1; Vb.Le(@x, @z); @x = Vb.Add(@x, 1)) {
                        // Normalized rule line 240
                        if (Vb.Bool(Vb.Eq(Vb.Mid(@str4, @x, 1), "-"))) {
                            break;
                        }
                        // Normalized rule line 242
                    }
                    // Normalized rule line 243
                    @lk = Vb.Cat("LK", Vb.Mid(@str4, Vb.Add(@x, 1), 1));
                    // Normalized rule line 244
                }
                // Normalized rule line 251
                if (Vb.Bool(Vb.Eq(PumpDatabaseSheet(@pump_type, @k), "ZE7"))) {
                    // Normalized rule line 252
                    @origpumptype = @pump_type;
                    // Normalized rule line 253
                    @pump_type = "ZE7";
                    // Normalized rule line 254
                }
                // Normalized rule line 257
                report("正在查询第 " + (Vb.Int(@k) + 1) + " 台泵：" + Vb.Text(@str4) + "（" + Vb.Text(@pump_type) + "）……");
                @databaseoffset = FindDatabaseOffset(@pump_type, @str4, @k);
                Book.Worksheets["Sheet1"].Range[Vb.Cat("F", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 24, false);
                // Normalized rule line 265
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 266
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 267
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 273
                if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range["$i$53"].Value2, "ASME"))) {
                    // Normalized rule line 275
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2 = Vb.Cat("DN", ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 2, false));
                    // Normalized rule line 276
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2 = Vb.Cat("DN", ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 3, false));
                    // Normalized rule line 278
                    if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "32"), 0))) {
                        // Normalized rule line 279
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2 = "DN15";
                        // Normalized rule line 281
                    } else {
                        // Normalized rule line 282
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 283
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AS", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 284
                    }
                    // Normalized rule line 287
                    if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "54"), 0))) {
                        // Normalized rule line 288
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2 = "DN15";
                        // Normalized rule line 290
                    } else {
                        // Normalized rule line 291
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 292
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AW", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 293
                    }
                    // Normalized rule line 296
                    if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "62"), 0))) {
                        // Normalized rule line 297
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2 = "DN15";
                        // Normalized rule line 299
                    } else {
                        // Normalized rule line 300
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 301
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AQ", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 302
                    }
                    // Normalized rule line 305
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 306
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AL", Vb.Add(40, @k))].Value2 = "DN20";
                        // Normalized rule line 308
                    } else {
                        // Normalized rule line 309
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AL", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 310
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 311
                    }
                    // Normalized rule line 313
                } else {
                    // Normalized rule line 314
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 4, false);
                    // Normalized rule line 315
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 5, false);
                    // Normalized rule line 317
                    if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "32"), 0))) {
                        // Normalized rule line 318
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2 = "1/2\"";
                        // Normalized rule line 320
                    } else {
                        // Normalized rule line 321
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 322
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AS", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 323
                    }
                    // Normalized rule line 326
                    if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "54"), 0))) {
                        // Normalized rule line 327
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2 = "1/2\"";
                        // Normalized rule line 329
                    } else {
                        // Normalized rule line 330
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 331
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AW", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 332
                    }
                    // Normalized rule line 335
                    if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "62"), 0))) {
                        // Normalized rule line 336
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2 = "1/2\"";
                        // Normalized rule line 338
                    } else {
                        // Normalized rule line 339
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 340
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AQ", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 341
                    }
                    // Normalized rule line 345
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 346
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AL", Vb.Add(40, @k))].Value2 = "3/4\"";
                        // Normalized rule line 348
                    } else {
                        // Normalized rule line 349
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AL", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 350
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = "NO";
                        // Normalized rule line 351
                    }
                    // Normalized rule line 354
                }
                // Normalized rule line 357
                if (Vb.Bool(Vb.Ne(@origpumptype, ""))) {
                    @pump_type = @origpumptype;
                }
                // Normalized rule line 360
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 361
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 362
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 369
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 370
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BS", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(@lk, Book.Worksheets["Data1"].Range["$g$65:$o$71"], 2, false);
                    // Normalized rule line 371
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BT", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(@lk, Book.Worksheets["Data1"].Range["$g$65:$o$71"], 3, false);
                    // Normalized rule line 372
                } else {
                    // Normalized rule line 373
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BS", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 374
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BT", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 375
                }
                // Normalized rule line 378
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 379
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BU", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(@lk, Book.Worksheets["Data1"].Range["$g$65:$o$71"], 4, false);
                    // Normalized rule line 380
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BV", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(@lk, Book.Worksheets["Data1"].Range["$g$65:$o$71"], 5, false);
                    // Normalized rule line 381
                } else {
                    // Normalized rule line 382
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BU", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 383
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BV", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 384
                }
                // Normalized rule line 388
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 389
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BW", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(@lk, Book.Worksheets["Data1"].Range["$g$65:$o$71"], 6, false);
                    // Normalized rule line 390
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BX", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(@lk, Book.Worksheets["Data1"].Range["$g$65:$o$71"], 7, false);
                    // Normalized rule line 391
                } else {
                    // Normalized rule line 392
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BW", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 393
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BX", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 394
                }
                // Normalized rule line 400
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 401
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 402
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 410
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                    // Normalized rule line 414
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AX", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 4, false);
                    // Normalized rule line 415
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AY", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 5, false);
                    // Normalized rule line 416
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AZ", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 6, false);
                    // Normalized rule line 420
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BA", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 14, false);
                    // Normalized rule line 421
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BB", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 15, false);
                    // Normalized rule line 422
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BC", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 16, false);
                    // Normalized rule line 425
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BK", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 20, false);
                    // Normalized rule line 426
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BL", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 21, false);
                    // Normalized rule line 427
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BM", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 22, false);
                    // Normalized rule line 430
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BN", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 20, false);
                    // Normalized rule line 431
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BO", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 21, false);
                    // Normalized rule line 432
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BP", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 22, false);
                    // Normalized rule line 434
                } else {
                    // Normalized rule line 435
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AX", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 4, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 436
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AY", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 5, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 437
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AZ", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 6, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 441
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BA", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 14, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 442
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BB", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 15, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 443
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BC", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 16, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 446
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BK", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 20, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 447
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BL", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 21, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 448
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BM", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("S", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 22, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 451
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BN", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 20, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 452
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BO", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 21, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 453
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BP", Vb.Add(40, @k))].Value2 = Vb.Mul(ExcelApp.WorksheetFunction.HLookup(Book.Worksheets["Sheet1"].Range[Vb.Cat("T", Vb.Add(40, @k))].Value2, Book.Worksheets["nozzle-load"].Range["$b$5:$AD$27"], 22, false), Book.Worksheets["input"].Range["$i$56"].Value2);
                    // Normalized rule line 497
                }
                // Normalized rule line 502
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 503
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 504
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 508
                if (Vb.Bool(Vb.Eq(PumpDatabaseSheet(@pump_type, @k), "ZE7"))) {
                    // Normalized rule line 509
                    @origpumptype = @pump_type;
                    // Normalized rule line 510
                    @pump_type = "ZE7";
                    // Normalized rule line 511
                }
                // Normalized rule line 515
                if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range[Vb.Cat("AC", Vb.Add(21, @k))].Value2, ""))) {
                    // Normalized rule line 517
                    if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range["$i$53"].Value2, "ASME"))) {
                        // Normalized rule line 519
                        if (Vb.Bool(Vb.And(Vb.And(Vb.And(Vb.Ne(@pump_type, "ZA"), Vb.Ne(@pump_type, "XA")), Vb.Ne(@pump_type, "XE")), Vb.Ne(@pump_type, "CZ")))) {
                            // Normalized rule line 522
                            if (Vb.Bool(Vb.Lt(ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 2, false), 80))) {
                                // Normalized rule line 523
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "DN15";
                                // Normalized rule line 524
                            } else {
                                // Normalized rule line 525
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "DN20";
                                // Normalized rule line 526
                            }
                            // Normalized rule line 528
                        } else {
                            // Normalized rule line 530
                            if (Vb.Bool(Vb.Lt(ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 2, false), 150))) {
                                // Normalized rule line 531
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "DN15";
                                // Normalized rule line 532
                            } else {
                                // Normalized rule line 533
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "DN25";
                                // Normalized rule line 534
                            }
                            // Normalized rule line 536
                        }
                        // Normalized rule line 539
                    } else {
                        // Normalized rule line 541
                        if (Vb.Bool(Vb.And(Vb.And(Vb.And(Vb.Ne(@pump_type, "ZA"), Vb.Ne(@pump_type, "XA")), Vb.Ne(@pump_type, "XE")), Vb.Ne(@pump_type, "CZ")))) {
                            // Normalized rule line 543
                            if (Vb.Bool(Vb.Or(Vb.Eq(Vb.Len(ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 4, false)), 3), Vb.Ge(Vb.Mid(ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 4, false), 1, 1), 3)))) {
                                // Normalized rule line 544
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "3/4\"";
                                // Normalized rule line 545
                            } else {
                                // Normalized rule line 546
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "1/2\"";
                                // Normalized rule line 547
                            }
                            // Normalized rule line 549
                        } else {
                            // Normalized rule line 550
                            if (Vb.Bool(Vb.Or(Vb.Eq(Vb.Len(ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 4, false)), 3), Vb.Ge(Vb.Mid(ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@pump_type].Range["$a$31:$bb$2000"], 4, false), 1, 1), 6)))) {
                                // Normalized rule line 551
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "1\"";
                                // Normalized rule line 552
                            } else {
                                // Normalized rule line 553
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = "1/2\"";
                                // Normalized rule line 554
                            }
                            // Normalized rule line 557
                        }
                        // Normalized rule line 559
                    }
                    // Normalized rule line 561
                } else {
                    // Normalized rule line 562
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("V", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("AC", Vb.Add(21, @k))].Value2;
                    // Normalized rule line 563
                }
                // Normalized rule line 565
                if (Vb.Bool(Vb.Ne(@origpumptype, ""))) {
                    @pump_type = @origpumptype;
                }
                // Normalized rule line 568
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 569
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 570
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 578
                if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range["$i$53"].Value2, "ASME"))) {
                    // Normalized rule line 580
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "ASME B16.5 150# RF";
                    }
                    // Normalized rule line 581
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "ASME B16.5 300#  RF";
                    }
                    // Normalized rule line 582
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "ASME B16.5 600#  RF";
                    }
                    // Normalized rule line 584
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "ASME B16.5 150#  RF";
                    }
                    // Normalized rule line 585
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "ASME B16.5 300#  RF";
                    }
                    // Normalized rule line 586
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "ASME B16.5 600#  RF";
                    }
                    // Normalized rule line 589
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 591
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = "ASME B16.5 150#  RF";
                        // Normalized rule line 593
                    } else {
                        // Normalized rule line 594
                    }
                    // Normalized rule line 596
                } else if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range["$i$53"].Value2, "HG20615"))) {
                    // Normalized rule line 598
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "HG20615 PN2.0 RF";
                    }
                    // Normalized rule line 599
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "HG20615 PN5.0 RF";
                    }
                    // Normalized rule line 600
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "HG20615 PN11.0 RF";
                    }
                    // Normalized rule line 602
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "HG20615 PN2.0 RF";
                    }
                    // Normalized rule line 603
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "HG20615 PN5.0 RF";
                    }
                    // Normalized rule line 604
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "HG20615 PN11.0 RF";
                    }
                    // Normalized rule line 607
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 609
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = "HG20615 PN2.0 RF";
                        // Normalized rule line 611
                    } else {
                        // Normalized rule line 612
                    }
                    // Normalized rule line 614
                } else if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range["$i$53"].Value2, "HG20592"))) {
                    // Normalized rule line 616
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "HG20592 PN2.5 RF";
                    }
                    // Normalized rule line 617
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "HG20592 PN4.0 RF";
                    }
                    // Normalized rule line 618
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "HG20615 PN11.0 RF";
                    }
                    // Normalized rule line 620
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "HG20592 PN2.5 RF";
                    }
                    // Normalized rule line 621
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "HG20592 PN4.0 RF";
                    }
                    // Normalized rule line 622
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "HG20615 PN11.0 RF";
                    }
                    // Normalized rule line 625
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 627
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = "HG20592 PN2.5 RF";
                        // Normalized rule line 629
                    } else {
                        // Normalized rule line 630
                    }
                    // Normalized rule line 632
                } else if (Vb.Bool(Vb.Eq(Book.Worksheets["input"].Range["$i$53"].Value2, "SH3406"))) {
                    // Normalized rule line 634
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "SH3406 PN2.0 RF";
                    }
                    // Normalized rule line 635
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "SH3406 PN5.0 RF";
                    }
                    // Normalized rule line 636
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = "SH3406 PN10.0 RF";
                    }
                    // Normalized rule line 638
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "SH3406 PN2.0 RF";
                    }
                    // Normalized rule line 639
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "SH3406 PN5.0 RF";
                    }
                    // Normalized rule line 640
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = "SH3406 PN10.0 RF";
                    }
                    // Normalized rule line 643
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 645
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = "SH3406 PN2.0 RF";
                        // Normalized rule line 647
                    } else {
                        // Normalized rule line 648
                    }
                    // Normalized rule line 650
                } else {
                    // Normalized rule line 652
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$12"].Value2), " RF");
                    }
                    // Normalized rule line 653
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$13"].Value2), " RF");
                    }
                    // Normalized rule line 654
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$14"].Value2), " RF");
                    }
                    // Normalized rule line 656
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "XA")), Vb.Eq(@pump_type, "CZ")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$12"].Value2), " RF");
                    }
                    // Normalized rule line 657
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZE"), Vb.Eq(@pump_type, "XE")))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$13"].Value2), " RF");
                    }
                    // Normalized rule line 658
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZF"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("W", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$14"].Value2), " RF");
                    }
                    // Normalized rule line 661
                    if (Vb.Bool(Vb.And(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, ""), Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("Z", Vb.Add(21, @k))].Value2, "无")))) {
                        // Normalized rule line 663
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AO", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Cat("GB9113.1-2000 ", Book.Worksheets["Sheet1"].Range["$c$12"].Value2), " RF");
                        // Normalized rule line 665
                    } else {
                        // Normalized rule line 666
                    }
                    // Normalized rule line 668
                }
                // Normalized rule line 675
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 676
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 677
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 683
                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "32"), 0))) {
                    // Normalized rule line 685
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AS", Vb.Add(40, @k))].Value2 = Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2;
                    // Normalized rule line 687
                } else {
                    // Normalized rule line 688
                }
                // Normalized rule line 690
                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "54"), 0))) {
                    // Normalized rule line 692
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AW", Vb.Add(40, @k))].Value2 = Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2;
                    // Normalized rule line 694
                } else {
                    // Normalized rule line 695
                }
                // Normalized rule line 697
                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "62"), 0))) {
                    // Normalized rule line 699
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("AQ", Vb.Add(40, @k))].Value2 = Book.Worksheets["Sheet1"].Range[Vb.Cat("U", Vb.Add(40, @k))].Value2;
                    // Normalized rule line 701
                } else {
                    // Normalized rule line 702
                }
                // Normalized rule line 708
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 709
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 710
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 730
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 731
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = "F";
                    // Normalized rule line 732
                } else {
                    // Normalized rule line 733
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Book.Worksheets["Sheet1"].Range[Vb.Cat("AR", Vb.Add(40, @k))].Value2;
                    // Normalized rule line 734
                }
                // Normalized rule line 736
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 737
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2, "B");
                    // Normalized rule line 738
                } else {
                    // Normalized rule line 739
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2, Book.Worksheets["Sheet1"].Range[Vb.Cat("AV", Vb.Add(40, @k))].Value2);
                    // Normalized rule line 740
                }
                // Normalized rule line 742
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 743
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2, "Q");
                    // Normalized rule line 744
                } else {
                    // Normalized rule line 745
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2, Book.Worksheets["Sheet1"].Range[Vb.Cat("AP", Vb.Add(40, @k))].Value2);
                    // Normalized rule line 746
                }
                // Normalized rule line 748
                if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("AL", Vb.Add(40, @k))].Value2, "NO"))) {
                    // Normalized rule line 749
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2, "C");
                    // Normalized rule line 750
                } else {
                    // Normalized rule line 751
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2 = Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("BD", Vb.Add(40, @k))].Value2, Book.Worksheets["Sheet1"].Range[Vb.Cat("AL", Vb.Add(40, @k))].Value2);
                    // Normalized rule line 752
                }
                // Normalized rule line 758
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 759
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 760
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 767
                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "21"), 0))) {
                    // Normalized rule line 768
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BE", Vb.Add(40, @k))].Value2 = "Y";
                    // Normalized rule line 769
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BI", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 770
                } else if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "23"), 0))) {
                    // Normalized rule line 771
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BE", Vb.Add(40, @k))].Value2 = "Y";
                    // Normalized rule line 773
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BI", Vb.Add(40, @k))].Value2 = "L";
                    // Normalized rule line 774
                } else if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "41"), 0))) {
                    // Normalized rule line 775
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BE", Vb.Add(40, @k))].Value2 = "Y";
                    // Normalized rule line 776
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BI", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 777
                } else {
                    // Normalized rule line 778
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BE", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 779
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BI", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 780
                }
                // Normalized rule line 782
                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "52"), 0))) {
                    // Normalized rule line 783
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BF", Vb.Add(40, @k))].Value2 = "Y";
                    // Normalized rule line 784
                } else if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("Y", Vb.Add(21, @k))].Value2, "53"), 0))) {
                    // Normalized rule line 785
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BF", Vb.Add(40, @k))].Value2 = "Y";
                    // Normalized rule line 787
                } else {
                    // Normalized rule line 788
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("BF", Vb.Add(40, @k))].Value2 = "N";
                    // Normalized rule line 789
                }
                // Normalized rule line 797
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 798
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 799
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 808
                if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2, ""))) {
                    // Normalized rule line 809
                    if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("B", Vb.Add(40, @k))].Value2, ""))) {
                        // Normalized rule line 810
                        @str3 = Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("B", Vb.Add(40, @k))].Value2, "_"), @str2new), "_"), Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2), "_"), Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2), "_Ver"), Book.Worksheets["Sheet1"].Range["$c$3"].Value2);
                        // Normalized rule line 811
                    } else {
                        // Normalized rule line 812
                        @str3 = Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(@str2new, "_"), Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2), "_"), Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2), "_Ver"), Book.Worksheets["Sheet1"].Range["$c$3"].Value2);
                        // Normalized rule line 813
                    }
                    // Normalized rule line 818
                } else {
                    // Normalized rule line 819
                    if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("B", Vb.Add(40, @k))].Value2, ""))) {
                        // Normalized rule line 820
                        @str3 = Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Book.Worksheets["Sheet1"].Range[Vb.Cat("B", Vb.Add(40, @k))].Value2, "_"), @str2), "_"), Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2), "_"), Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2), "_Ver"), Book.Worksheets["Sheet1"].Range["$c$3"].Value2);
                        // Normalized rule line 821
                    } else {
                        // Normalized rule line 822
                        @str3 = Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(Vb.Cat(@str2, "_"), Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2), "_"), Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2), "_Ver"), Book.Worksheets["Sheet1"].Range["$c$3"].Value2);
                        // Normalized rule line 823
                    }
                    // Normalized rule line 825
                }
                // Normalized rule line 833
                @j = Vb.Len(@str3);
                // Normalized rule line 834
                @name = "";
                // Normalized rule line 835
                for (@x = 1; Vb.Le(@x, @j); @x = Vb.Add(@x, 1)) {
                    // Normalized rule line 837
                    @str1 = Vb.Mid(@str3, @x, 1);
                    // Normalized rule line 840
                    if (Vb.Bool(Vb.And(Vb.And(Vb.And(Vb.And(Vb.And(Vb.And(Vb.Ne(@str1, "/"), Vb.Ne(@str1, "\\")), Vb.Ne(@str1, " ")), Vb.Ne(@str1, ")")), Vb.Ne(@str1, "(")), Vb.Ne(@str1, ",")), Vb.Ne(@str1, "，")))) {
                        // Normalized rule line 841
                        @name = Vb.Cat(@name, @str1);
                        // Normalized rule line 842
                    } else {
                        // Normalized rule line 843
                        @str1 = "";
                        // Normalized rule line 844
                        @name = Vb.Cat(@name, @str1);
                        // Normalized rule line 845
                    }
                    // Normalized rule line 846
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("P", Vb.Add(40, @k))].Value2 = Vb.Cat("GA_", @name);
                    // Normalized rule line 847
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("Q", Vb.Add(40, @k))].Value2 = Vb.Cat("PID_", @name);
                    // Normalized rule line 848
                    Book.Worksheets["Sheet1"].Range[Vb.Cat("R", Vb.Add(40, @k))].Value2 = Vb.Cat("Sec&List_", @name);
                    // Normalized rule line 850
                }
                // Normalized rule line 856
                @gadatabasesheet = PumpDatabaseSheet(@pump_type, @k);
                // Normalized rule line 857
                @counter = Vb.Sub(Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("A", ExcelApp.Rows.Count)].End[-4162].Row, 31);
                // Normalized rule line 866
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 867
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 868
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 881
                // Normalized rule line 882
                @databaseoffset = FindDatabaseOffset(@gadatabasesheet, @str4, @k);
                // Normalized rule line 883
                for (@x = @databaseoffset; Vb.Le(@x, @databaseoffset); @x = Vb.Add(@x, 1)) {
                    // Normalized rule line 885
                    @counter2 = 0;
                    // Normalized rule line 886
                    @other_weight = 20;
                    // Normalized rule line 887
                    if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("V", Vb.Add(21, @k))].Value2, "380V"))) {
                        // Normalized rule line 889
                        if (Vb.Bool(Vb.And(Vb.Eq(Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("A", Vb.Add(32, @x))].Value2, @str4), Vb.Eq(Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("F", Vb.Add(32, @x))].Value2, Vb.Cat(Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2, "(6KV)"))))) {
                            // Normalized rule line 891
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("G", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("y", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 893
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("H", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("z", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 896
                            for (@y = 0; Vb.Le(@y, 6); @y = Vb.Add(@y, 1)) {
                                // Normalized rule line 898
                                @str5 = Book.Worksheets["Data1"].Range[Vb.Cat("A", Vb.Add(2, @y))].Value2;
                                // Normalized rule line 899
                                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2, @str5), 0))) {
                                    // Normalized rule line 900
                                    @counter2 = Vb.Add(@counter2, 1);
                                    // Normalized rule line 902
                                } else {
                                    // Normalized rule line 903
                                }
                                // Normalized rule line 907
                            }
                            // Normalized rule line 908
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("I", Vb.Add(40, @k))].Value2 = Vb.Add(@other_weight, Vb.Mul(100, @counter2));
                            // Normalized rule line 916
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("X", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("g", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 918
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("Y", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("h", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 920
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("Z", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("i", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 922
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AA", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("j", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 924
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AB", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("k", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 926
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AC", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("l", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 927
                            if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2, ""))) {
                                // Normalized rule line 928
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("CN", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2;
                                // Normalized rule line 929
                            }
                            // Normalized rule line 932
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AD", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("m", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 934
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AE", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("n", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 936
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AF", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("o", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 938
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AG", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("p", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 940
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AH", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("q", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 942
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AI", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("r", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 944
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AJ", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("s", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 946
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AM", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("v", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 948
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AN", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("w", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 950
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AT", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("ac", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 952
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AU", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("ad", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 954
                            if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "CZ")))) {
                                // Normalized rule line 955
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("BQ", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("ae", Vb.Add(@x, 32))].Value2;
                                // Normalized rule line 956
                            } else {
                                // Normalized rule line 957
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("BQ", Vb.Add(40, @k))].Value2 = "N";
                                // Normalized rule line 958
                            }
                            // Normalized rule line 960
                        } else {
                            // Normalized rule line 962
                        }
                        // Normalized rule line 964
                    } else {
                        // Normalized rule line 966
                        if (Vb.Bool(Vb.And(Vb.Eq(Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("A", Vb.Add(32, @x))].Value2, @str4), Vb.Eq(Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("F", Vb.Add(32, @x))].Value2, Book.Worksheets["input"].Range[Vb.Cat("U", Vb.Add(21, @k))].Value2)))) {
                            // Normalized rule line 967
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("G", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("y", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 968
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("H", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("z", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 971
                            for (@y = 0; Vb.Le(@y, 6); @y = Vb.Add(@y, 1)) {
                                // Normalized rule line 973
                                @str5 = Book.Worksheets["Data1"].Range[Vb.Cat("A", Vb.Add(2, @y))].Value2;
                                // Normalized rule line 974
                                if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["Sheet1"].Range[Vb.Cat("E", Vb.Add(40, @k))].Value2, @str5), 0))) {
                                    // Normalized rule line 975
                                    @counter2 = Vb.Add(@counter2, 1);
                                    // Normalized rule line 977
                                } else {
                                    // Normalized rule line 978
                                }
                                // Normalized rule line 982
                            }
                            // Normalized rule line 983
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("I", Vb.Add(40, @k))].Value2 = Vb.Add(@other_weight, Vb.Mul(100, @counter2));
                            // Normalized rule line 987
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("X", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("g", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 988
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("Y", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("h", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 989
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("Z", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("i", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 990
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AA", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("j", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 991
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AB", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("k", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 992
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AC", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("l", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 994
                            if (Vb.Bool(Vb.Ne(Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2, ""))) {
                                // Normalized rule line 995
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("CN", Vb.Add(40, @k))].Value2 = Book.Worksheets["input"].Range[Vb.Cat("AB", Vb.Add(21, @k))].Value2;
                                // Normalized rule line 996
                            }
                            // Normalized rule line 999
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AD", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("m", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1000
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AE", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("n", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1001
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AF", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("o", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1002
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AG", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("p", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1003
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AH", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("q", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1004
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AI", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("r", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1005
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AJ", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("s", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1006
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AM", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("v", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1007
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AN", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("w", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1008
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AT", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("ac", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1009
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AU", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("ad", Vb.Add(@x, 32))].Value2;
                            // Normalized rule line 1010
                            if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "CZ")))) {
                                // Normalized rule line 1011
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("BQ", Vb.Add(40, @k))].Value2 = Book.Worksheets[@gadatabasesheet].Range[Vb.Cat("ae", Vb.Add(@x, 32))].Value2;
                                // Normalized rule line 1012
                            } else {
                                // Normalized rule line 1013
                                Book.Worksheets["Sheet1"].Range[Vb.Cat("BQ", Vb.Add(40, @k))].Value2 = "N";
                                // Normalized rule line 1014
                            }
                            // Normalized rule line 1017
                        } else {
                            // Normalized rule line 1019
                        }
                        // Normalized rule line 1023
                    }
                    // Normalized rule line 1029
                    if (Vb.Bool(Vb.Ne(@pump_type, "ZE"))) {
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@gadatabasesheet].Range["$a$31:$bb$2000"], 20, false);
                    }
                    // Normalized rule line 1031
                    if (Vb.Bool(Vb.Eq(@pump_type, "ZE"))) {
                        // Normalized rule line 1032
                        if (Vb.Bool(Vb.And(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("E", Vb.Add(21, @k))].Value2, "8"), 0), Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("AA", Vb.Add(21, @k))].Value2, "F")))) {
                            // Normalized rule line 1033
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@gadatabasesheet].Range["$a$31:$bb$2000"], 21, false);
                            // Normalized rule line 1036
                        } else if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("E", Vb.Add(21, @k))].Value2, "0th"), 0))) {
                            // Normalized rule line 1037
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@gadatabasesheet].Range["$a$31:$bb$2000"], 21, false);
                            // Normalized rule line 1039
                        } else if (Vb.Bool(Vb.Ne(Vb.Instr(Book.Worksheets["input"].Range[Vb.Cat("E", Vb.Add(21, @k))].Value2, "7"), 0))) {
                            // Normalized rule line 1040
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@gadatabasesheet].Range["$a$31:$bb$2000"], 20, false);
                            // Normalized rule line 1042
                        } else {
                            // Normalized rule line 1043
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup(@str4, Book.Worksheets[@gadatabasesheet].Range["$a$31:$bb$2000"], 20, false);
                            // Normalized rule line 1044
                        }
                        // Normalized rule line 1045
                    }
                    // Normalized rule line 1051
                    if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("BE", Vb.Add(40, @k))].Value2, "N"))) {
                        // Normalized rule line 1053
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2 = Vb.Sub(Vb.Add(Book.Worksheets["Sheet1"].Range[Vb.Cat("AF", Vb.Add(40, @k))].Value2, Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2), 200);
                        // Normalized rule line 1056
                        if (Vb.Bool(Vb.Ne(Vb.Mid(Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2, Vb.Len(Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2), 1), 5))) {
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Mid(Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2, 1, Vb.Sub(Vb.Len(Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2), 1)), "0");
                        }
                        // Normalized rule line 1058
                    } else {
                        // Normalized rule line 1059
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("BG", Vb.Add(40, @k))].Value2 = "N";
                        // Normalized rule line 1060
                    }
                    // Normalized rule line 1062
                    if (Vb.Bool(Vb.Ne(Book.Worksheets["Sheet1"].Range[Vb.Cat("BF", Vb.Add(40, @k))].Value2, "N"))) {
                        // Normalized rule line 1063
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2 = Vb.Sub(Vb.Add(Book.Worksheets["Sheet1"].Range[Vb.Cat("AF", Vb.Add(40, @k))].Value2, Book.Worksheets["Sheet1"].Range[Vb.Cat("AK", Vb.Add(40, @k))].Value2), 100);
                        // Normalized rule line 1065
                        if (Vb.Bool(Vb.Ne(Vb.Mid(Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2, Vb.Len(Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2), 1), 5))) {
                            Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2 = Vb.Cat(Vb.Mid(Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2, 1, Vb.Sub(Vb.Len(Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2), 1)), "0");
                        }
                        // Normalized rule line 1069
                    } else {
                        // Normalized rule line 1071
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("BH", Vb.Add(40, @k))].Value2 = "N";
                        // Normalized rule line 1073
                    }
                    // Normalized rule line 1076
                    if (Vb.Bool(Vb.Or(Vb.Eq(@pump_type, "ZA"), Vb.Eq(@pump_type, "CZ")))) {
                        // Normalized rule line 1078
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("BJ", Vb.Add(40, @k))].Value2 = "1\"";
                        // Normalized rule line 1079
                    } else {
                        // Normalized rule line 1081
                        Book.Worksheets["Sheet1"].Range[Vb.Cat("BJ", Vb.Add(40, @k))].Value2 = "2\"";
                        // Normalized rule line 1083
                    }
                    // Normalized rule line 1091
                }
                // Normalized rule line 1097
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 1098
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 1099
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 1105
                CalculateSectionNames(@k);
                // Normalized rule line 1108
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 1109
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 1110
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 1119
                @progress = Vb.Mul(5, @counterprogress);
                // Normalized rule line 1120
                @counterprogress = Vb.Add(@counterprogress, 1);
                // Normalized rule line 1121
                @pctdone = Vb.Div(@progress, @countersum);
                // Normalized rule line 1126
            }
        }
    }
}
