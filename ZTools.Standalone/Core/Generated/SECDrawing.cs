// Ported from migration/legacy; see tools/port_vba.py and migration/README.md.
using System;
namespace ZTools.Engine {
    internal sealed partial class SubmissionEngine {
        public void BuildSEC() {
            dynamic @anchor = "";
            dynamic @attrib = null;
            dynamic @bearing = null;
            dynamic @blockattributes = null;
            dynamic @bold = false;
            
            dynamic @charset = 0;
            dynamic @cool = "";
            dynamic @count = null;
            dynamic @counter = 0;
            dynamic @drainza = null;
            dynamic @exchanger = null;
            dynamic @gongsimingcheng = "";
            dynamic @i = null;
            dynamic @index = null;
            dynamic @italic = false;
            dynamic @j = null;
            dynamic @jiekou = null;
            dynamic @k = null;
            dynamic @last = "";
            dynamic @lk = "";
            dynamic @m = 0;
            dynamic @mscale = null;
            dynamic @mtext1 = null;
            dynamic @mtext2 = null;
            dynamic @mtext3 = null;
            dynamic @mtext4 = null;
            dynamic @myobj = null;
            dynamic @myset = null;
            dynamic @n = 0;
            dynamic @nam = null;
            dynamic @newpath = "";
            dynamic @newpumptype = "";
            dynamic @nordia = "";
            dynamic @obj = null;
            dynamic @objdoc = null;
            dynamic @objlayout = null;
            dynamic @objold = null;
            dynamic @objplot = null;
            dynamic @p = new double[3];
            dynamic @pathname = "";
            dynamic @pitchandfamily = 0;
            dynamic @ptmax = null;
            dynamic @ptmin = null;
            dynamic @pump = null;
            dynamic @q = new double[3];
            dynamic @r = new double[3];
            dynamic @reservoir = "";
            dynamic @s = new double[3];
            dynamic @savename = null;
            dynamic @shebeimingcheng = "";
            dynamic @side = null;
            dynamic @size = null;
            dynamic @ss = null;
            dynamic @stz = null;
            dynamic @t = new double[3];
            dynamic @tagname = null;
            dynamic @tempstring = null;
            dynamic @th = "";
            dynamic @typeface = "";
            dynamic @u = new double[3];
            dynamic @v = new double[3];
            dynamic @ver = "";
            dynamic @w = new double[3];
            dynamic @x = new double[3];
            dynamic @xiangmumingcheng = null;
            dynamic @xlapp = null;
            dynamic @xlbook = null;
            dynamic @xlsheet = null;
            dynamic @xlsheet1 = null;
            dynamic @xlsheet2 = null;
            dynamic @xlsheet3 = null;
            dynamic @yonghudingdanhao = null;
            dynamic @yonghumingcheng = null;
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
            // Normalized rule line 12
            // Normalized rule line 13
            // Normalized rule line 14
            // Normalized rule line 15
            // Normalized rule line 16
            // Normalized rule line 17
            // Normalized rule line 18
            // Normalized rule line 19
            // Normalized rule line 20
            // Normalized rule line 21
            // Normalized rule line 22
            // Normalized rule line 23
            // Normalized rule line 24
            // Normalized rule line 25
            // Normalized rule line 26
            // Normalized rule line 27
            // Normalized rule line 28
            // Normalized rule line 29
            // Normalized rule line 30
            // Normalized rule line 31
            // Normalized rule line 32
            // Normalized rule line 33
            // Normalized rule line 34
            // Normalized rule line 35
            // Normalized rule line 36
            // Normalized rule line 37
            // Normalized rule line 38
            // Normalized rule line 39
            // Normalized rule line 40
            // Normalized rule line 41
            // Normalized rule line 42
            // Normalized rule line 43
            // Normalized rule line 44
            // Normalized rule line 49
            @xlbook = Book;
            // Normalized rule line 50
            @xlsheet = Book.Worksheets["Sheet1"];
            // Normalized rule line 51
            @xlsheet1 = Book.Worksheets["Sec-List"];
            // Normalized rule line 52
            @xlsheet2 = Book.Worksheets["input"];
            // Normalized rule line 53
            @xlsheet3 = Book.Worksheets["CZ-BearingSize"];
            // Normalized rule line 54
            @i = Vb.Val(@xlsheet.Range[Vb.Cat("C", 19)].Value2);
            // Normalized rule line 55
            @savename = @xlsheet.Range[Vb.Cat("R", @i)].Value2;
            // Normalized rule line 56
            @shebeimingcheng = @xlsheet.Range[Vb.Cat("D", @i)].Value2;
            // Normalized rule line 57
            @yonghudingdanhao = @xlsheet.Range[Vb.Cat("C", 9)].Value2;
            // Normalized rule line 58
            @xiangmumingcheng = @xlsheet.Range[Vb.Cat("C", 8)].Value2;
            // Normalized rule line 59
            @yonghumingcheng = @xlsheet.Range[Vb.Cat("C", 7)].Value2;
            // Normalized rule line 60
            @gongsimingcheng = @xlsheet2.Range[Vb.Cat("K", 10)].Value2;
            // Normalized rule line 61
            if (Vb.Bool(Vb.Ne(@xlsheet.Range[Vb.Cat("CN", @i)].Value2, ""))) {
                // Normalized rule line 62
                @newpumptype = Vb.Replace(@xlsheet.Range[Vb.Cat("C", @i)].Value2, Vb.Left(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 2), @xlsheet.Range[Vb.Cat("CN", @i)].Value2);
                // Normalized rule line 63
            } else {
                // Normalized rule line 64
                @newpumptype = @xlsheet.Range[Vb.Cat("C", @i)].Value2;
                // Normalized rule line 65
            }
            // Normalized rule line 66
            @j = Vb.Val(@xlsheet.Range[Vb.Cat("AU", @i)].Value2);
            // Normalized rule line 67
            @jiekou = @xlsheet.Range[Vb.Cat("BD", @i)].Value2;
            // Normalized rule line 68
            @size = @xlsheet.Range[Vb.Cat("C", @i)].Value2;
            // Normalized rule line 69
            @pump = Vb.Left(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 2);
            // Normalized rule line 70
            @nordia = Vb.Right(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 3);
            // Normalized rule line 71
            @anchor = Vb.Right(@xlsheet.Range[Vb.Cat("AT", @i)].Value2, 7);
            // Normalized rule line 72
            @exchanger = @xlsheet.Range[Vb.Cat("BE", @i)].Value2;
            // Normalized rule line 73
            @reservoir = @xlsheet.Range[Vb.Cat("BF", @i)].Value2;
            // Normalized rule line 74
            @side = @xlsheet.Range[Vb.Cat("BI", @i)].Value2;
            // Normalized rule line 75
            @cool = @xlsheet.Range[Vb.Cat("AL", @i)].Value2;
            // Normalized rule line 76
            @th = Vb.Right(Vb.Left(@xlsheet.Range[Vb.Cat("CA", @i)].Value2, 3), 1);
            // Normalized rule line 77
            @bearing = @xlsheet.Range[Vb.Cat("CB", @i)].Value2;
            // Normalized rule line 78
            @p[Vb.Int(0)] = 0;
            // Normalized rule line 79
            @p[Vb.Int(1)] = 0;
            // Normalized rule line 80
            @p[Vb.Int(2)] = 0;
            // Normalized rule line 81
            @q[Vb.Int(0)] = 420;
            // Normalized rule line 82
            @q[Vb.Int(1)] = 297;
            // Normalized rule line 83
            @q[Vb.Int(2)] = 0;
            // Normalized rule line 84
            @r[Vb.Int(0)] = 361;
            // Normalized rule line 85
            @v[Vb.Int(0)] = 361;
            // Normalized rule line 86
            @w[Vb.Int(0)] = 361;
            // Normalized rule line 87
            @x[Vb.Int(0)] = 361;
            // Normalized rule line 88
            @r[Vb.Int(1)] = 63;
            // Normalized rule line 89
            @v[Vb.Int(1)] = 67;
            // Normalized rule line 90
            @x[Vb.Int(1)] = 71.5;
            // Normalized rule line 91
            @w[Vb.Int(1)] = 53.9;
            // Normalized rule line 92
            @r[Vb.Int(2)] = 0;
            // Normalized rule line 93
            @v[Vb.Int(2)] = 0;
            // Normalized rule line 94
            @w[Vb.Int(2)] = 0;
            // Normalized rule line 95
            @x[Vb.Int(2)] = 0;
            // Normalized rule line 96
            @t[Vb.Int(0)] = 0;
            // Normalized rule line 97
            @t[Vb.Int(1)] = Vb.Neg(4.5);
            // Normalized rule line 98
            @t[Vb.Int(2)] = 0;
            // Normalized rule line 102
            @ver = @th;
            // Normalized rule line 104
            @ptmin = @p;
            // Normalized rule line 105
            @ptmax = @q;
            // Normalized rule line 106
            if (Vb.Bool(Vb.Or(Vb.Eq(@th, "A"), Vb.Eq(@pump, "ZF")))) {
                // Normalized rule line 107
                @ver = "";
                // Normalized rule line 108
            } else {
                // Normalized rule line 109
                if (Vb.Bool(Vb.And(Vb.Eq(@th, "1"), Vb.Eq(@pump, "ZE")))) {
                    // Normalized rule line 110
                    @ver = 10;
                    // Normalized rule line 111
                }
                // Normalized rule line 112
            }
            // Normalized rule line 114
            if (Vb.Bool(Vb.Eq(@pump, "CZ"))) {
                // Normalized rule line 115
                @lk = Vb.Right(Vb.Left(@xlsheet.Range[Vb.Cat("CD", @i)].Value2, 3), 1);
                // Normalized rule line 116
            } else {
                // Normalized rule line 117
                if (Vb.Bool(Vb.Eq(@xlsheet.Range[Vb.Cat("BZ", @i)].Value2, "Y"))) {
                    // Normalized rule line 118
                    @lk = Vb.Cat("_", "K6");
                    // Normalized rule line 119
                } else {
                    // Normalized rule line 120
                    @lk = "";
                    // Normalized rule line 121
                }
                // Normalized rule line 122
            }
            // Normalized rule line 137
            @mtext1 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@r, 48, @shebeimingcheng);
            // Normalized rule line 140
            dynamic case1 = Vb.Lenb(@shebeimingcheng);
            // Normalized rule line 141
            if (Vb.Bool(Vb.Lt(case1, 18))) {
                // Normalized rule line 142
                @mtext1.ScaleEntity(@r, 1);
                // Normalized rule line 143
                @mtext1.Height = 2.3;
                // Normalized rule line 144
                @mtext1.Width = 48;
                // Normalized rule line 145
                @mtext1.Update();
                // Normalized rule line 146
            } else if (Vb.Bool(Vb.And(Vb.Ge(case1, 18), Vb.Le(case1, 36)))) {
                // Normalized rule line 147
                @mtext1.ScaleEntity(@r, 0.85);
                // Normalized rule line 148
                @mtext1.Height = 2;
                // Normalized rule line 149
                @mtext1.Width = 55;
                // Normalized rule line 150
                @mtext1.Update();
                // Normalized rule line 151
            } else if (Vb.Bool(Vb.Gt(case1, 36))) {
                // Normalized rule line 152
                @mtext1.ScaleEntity(@r, 0.6);
                // Normalized rule line 153
                @mtext1.Height = 1.65;
                // Normalized rule line 154
                @mtext1.Width = 75;
                // Normalized rule line 155
                @mtext1.Update();
                // Normalized rule line 156
            }
            // Normalized rule line 159
            @mtext2 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@v, 48, @xiangmumingcheng);
            // Normalized rule line 162
            dynamic case2 = Vb.Lenb(@xiangmumingcheng);
            // Normalized rule line 163
            if (Vb.Bool(Vb.Lt(case2, 18))) {
                // Normalized rule line 164
                @mtext2.ScaleEntity(@v, 1);
                // Normalized rule line 165
                @mtext2.Height = 2.3;
                // Normalized rule line 166
                @mtext2.Width = 48;
                // Normalized rule line 167
                @mtext2.Update();
                // Normalized rule line 168
            } else if (Vb.Bool(Vb.And(Vb.Ge(case2, 18), Vb.Le(case2, 36)))) {
                // Normalized rule line 169
                @mtext2.ScaleEntity(@v, 0.85);
                // Normalized rule line 170
                @mtext2.Height = 2;
                // Normalized rule line 171
                @mtext2.Width = 55;
                // Normalized rule line 172
                @mtext2.Update();
                // Normalized rule line 173
            } else if (Vb.Bool(Vb.Gt(case2, 36))) {
                // Normalized rule line 174
                @mtext2.ScaleEntity(@v, 0.6);
                // Normalized rule line 175
                @mtext2.Height = 1.65;
                // Normalized rule line 176
                @mtext2.Width = 75;
                // Normalized rule line 177
                @mtext2.Update();
                // Normalized rule line 178
            }
            // Normalized rule line 180
            @mtext3 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@w, 48, @yonghudingdanhao);
            // Normalized rule line 183
            dynamic case3 = Vb.Lenb(@yonghudingdanhao);
            // Normalized rule line 184
            if (Vb.Bool(Vb.Lt(case3, 18))) {
                // Normalized rule line 185
                @mtext3.ScaleEntity(@w, 1);
                // Normalized rule line 186
                @mtext3.Height = 2.3;
                // Normalized rule line 187
                @mtext3.Width = 48;
                // Normalized rule line 188
                @mtext3.Update();
                // Normalized rule line 189
            } else if (Vb.Bool(Vb.And(Vb.Ge(case3, 18), Vb.Le(case3, 36)))) {
                // Normalized rule line 190
                @mtext3.ScaleEntity(@w, 0.85);
                // Normalized rule line 191
                @mtext3.Height = 2;
                // Normalized rule line 192
                @mtext3.Width = 55;
                // Normalized rule line 193
                @mtext3.Update();
                // Normalized rule line 194
            } else if (Vb.Bool(Vb.Gt(case3, 36))) {
                // Normalized rule line 195
                @mtext3.ScaleEntity(@w, 0.6);
                // Normalized rule line 196
                @mtext3.Height = 1.6;
                // Normalized rule line 197
                @mtext3.Width = 70;
                // Normalized rule line 198
                @mtext3.Update();
                // Normalized rule line 199
            }
            // Normalized rule line 201
            @mtext4 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@x, 48, @yonghumingcheng);
            // Normalized rule line 204
            dynamic case4 = Vb.Lenb(@yonghumingcheng);
            // Normalized rule line 205
            if (Vb.Bool(Vb.Lt(case4, 18))) {
                // Normalized rule line 206
                @mtext4.ScaleEntity(@x, 1);
                // Normalized rule line 207
                @mtext4.Height = 2.3;
                // Normalized rule line 208
                @mtext4.Width = 48;
                // Normalized rule line 209
                @mtext4.Update();
                // Normalized rule line 210
            } else if (Vb.Bool(Vb.And(Vb.Ge(case4, 18), Vb.Le(case4, 36)))) {
                // Normalized rule line 211
                @mtext4.ScaleEntity(@x, 0.85);
                // Normalized rule line 212
                @mtext4.Height = 2;
                // Normalized rule line 213
                @mtext4.Width = 55;
                // Normalized rule line 214
                @mtext4.Update();
                // Normalized rule line 215
            } else if (Vb.Bool(Vb.Gt(case4, 36))) {
                // Normalized rule line 216
                @mtext4.ScaleEntity(@x, 0.6);
                // Normalized rule line 217
                @mtext4.Height = 1.6;
                // Normalized rule line 218
                @mtext4.Width = 70;
                // Normalized rule line 219
                @mtext4.Update();
                // Normalized rule line 220
            }
            // Normalized rule line 223
            if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 18)].Value2), 1))) {
                // Normalized rule line 224
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("Drain", @pump), 1, 1, 1, 0);
                // Normalized rule line 225
            }
            // Normalized rule line 227
            if (Vb.Bool(Vb.And(Vb.And(Vb.Ne(@pump, "ZA"), Vb.Ne(@pump, "XA")), Vb.Ne(@pump, "CZ")))) {
                // Normalized rule line 228
                if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 17)].Value2), 1))) {
                    // Normalized rule line 229
                    Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "PUREOIL", 1, 1, 1, 0);
                    // Normalized rule line 230
                }
                // Normalized rule line 232
            }
            // Normalized rule line 235
            InsertSection(@p, @i);
            // Normalized rule line 236
            @count = Drawing.Layouts.Item("Model").Block.Count;
            // Normalized rule line 237
            for (@index = 0; Vb.Le(@index, Vb.Sub(@count, 1)); @index = Vb.Add(@index, 1)) {
                // Normalized rule line 239
                @tempstring = Drawing.Layouts.Item("Model").Block.Item(@index).ObjectName;
                // Normalized rule line 240
                if (Vb.Bool(Vb.Eq(@tempstring, "AcDbBlockReference"))) {
                    // Normalized rule line 241
                    if (Vb.Bool(true)) {
                        // Normalized rule line 243
                        @blockattributes = Drawing.Layouts.Item("Model").Block.Item(@index).GetAttributes();
                        // Normalized rule line 244
                        foreach (dynamic each5 in @blockattributes) {
                            @attrib = each5;
                            // Normalized rule line 245
                            @tagname = @attrib.TagString;
                            // Normalized rule line 246
                            dynamic case6 = Vb.Ucase(@tagname);
                            // Normalized rule line 257
                            if (Vb.Bool(Vb.Eq(case6, "SHEBEIWEIHAO"))) {
                                // Normalized rule line 259
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("B", @i)].Value2);
                                // Normalized rule line 263
                            } else if (Vb.Bool(Vb.Eq(case6, "DINGDANHAO_SULZER"))) {
                                // Normalized rule line 265
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 10)].Value2);
                                // Normalized rule line 266
                            } else if (Vb.Bool(Vb.Eq(case6, "CAILIAODAIMA"))) {
                                // Normalized rule line 268
                                @attrib.TextString = Vb.Text(@xlsheet1.Range[Vb.Cat("F", Vb.Sub(@i, 30))].Value2);
                                // Normalized rule line 271
                            } else if (Vb.Bool(Vb.Eq(case6, "TYPE_PUMP"))) {
                                // Normalized rule line 274
                                @attrib.TextString = Vb.Text(Vb.Cat("   ", @newpumptype));
                                // Normalized rule line 276
                            } else if (Vb.Bool(Vb.Eq(case6, "DATE"))) {
                                // Normalized rule line 278
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 6)].Text);
                                // Normalized rule line 279
                            } else if (Vb.Bool(Vb.Eq(case6, "DESIGN"))) {
                                // Normalized rule line 281
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 4)].Value2);
                                // Normalized rule line 282
                            } else if (Vb.Bool(Vb.Eq(case6, "CHECK"))) {
                                // Normalized rule line 284
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 5)].Value2);
                                // Normalized rule line 285
                            } else if (Vb.Bool(Vb.Eq(case6, "COMPANY"))) {
                                // Normalized rule line 287
                                @attrib.TextString = Vb.Text(@gongsimingcheng);
                                // Normalized rule line 289
                            }
                            // Normalized rule line 290
                        }
                        // Normalized rule line 291
                    }
                    // Normalized rule line 292
                }
                // Normalized rule line 293
            }
            // Normalized rule line 296
            for (@n = 1; Vb.Le(@n, 40); @n = Vb.Add(@n, 1)) {
                // Normalized rule line 298
                @count = Drawing.Layouts.Item("Model").Block.Count;
                // Normalized rule line 300
                for (@index = 0; Vb.Le(@index, Vb.Sub(@count, 1)); @index = Vb.Add(@index, 1)) {
                    // Normalized rule line 302
                    @tempstring = Drawing.Layouts.Item("Model").Block.Item(@index).ObjectName;
                    // Normalized rule line 303
                    if (Vb.Bool(Vb.Eq(@tempstring, "AcDbBlockReference"))) {
                        // Normalized rule line 304
                        if (Vb.Bool(true)) {
                            // Normalized rule line 306
                            @blockattributes = Drawing.Layouts.Item("Model").Block.Item(@index).GetAttributes();
                            // Normalized rule line 307
                            foreach (dynamic each7 in @blockattributes) {
                                @attrib = each7;
                                // Normalized rule line 308
                                @tagname = @attrib.TagString;
                                // Normalized rule line 309
                                dynamic case8 = Vb.Ucase(@tagname);
                                // Normalized rule line 311
                                if (Vb.Bool(Vb.Eq(case8, Vb.Cat("PN", @n)))) {
                                    // Normalized rule line 313
                                    @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("DB", Vb.Add(@n, 71))].Value2);
                                    // Normalized rule line 314
                                } else if (Vb.Bool(Vb.Eq(case8, Vb.Cat("DES", @n)))) {
                                    // Normalized rule line 316
                                    @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("DC", Vb.Add(@n, 71))].Value2);
                                    // Normalized rule line 317
                                } else if (Vb.Bool(Vb.Eq(case8, Vb.Cat("MAT", @n)))) {
                                    // Normalized rule line 319
                                    @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("DD", Vb.Add(@n, 71))].Value2);
                                    // Normalized rule line 320
                                } else if (Vb.Bool(Vb.Eq(case8, Vb.Cat("QTY", @n)))) {
                                    // Normalized rule line 322
                                    @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("DE", Vb.Add(@n, 71))].Value2);
                                    // Normalized rule line 324
                                }
                                // Normalized rule line 325
                            }
                            // Normalized rule line 326
                        }
                        // Normalized rule line 327
                    }
                    // Normalized rule line 328
                }
                // Normalized rule line 330
            }
        }
    }
}
