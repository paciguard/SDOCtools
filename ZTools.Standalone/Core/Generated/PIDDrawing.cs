// Ported from migration/legacy; see tools/port_vba.py and migration/README.md.
using System;
namespace ZTools.Engine {
    internal sealed partial class SubmissionEngine {
        public void BuildPID() {
            dynamic @anchor = "";
            dynamic @attrib = null;
            dynamic @bearing = null;
            dynamic @blockattributes = null;
            dynamic @bold = false;
            
            dynamic @charset = 0;
            dynamic @cool = "";
            dynamic @count = null;
            dynamic @counter = 0;
            dynamic @drainza = "";
            dynamic @exchanger = null;
            dynamic @i = null;
            dynamic @index = null;
            dynamic @italic = false;
            dynamic @j = null;
            dynamic @jiekou = null;
            dynamic @k = null;
            dynamic @last = "";
            dynamic @lk = null;
            dynamic @lp1 = null;
            dynamic @lp2 = null;
            dynamic @lp3 = null;
            dynamic @lp4 = null;
            dynamic @lp5 = null;
            dynamic @lp6 = "";
            dynamic @m = 0;
            dynamic @mscale = null;
            dynamic @mtext1 = null;
            dynamic @mtext2 = null;
            dynamic @mtext3 = null;
            dynamic @mtext4 = null;
            dynamic @myobj = null;
            dynamic @myset = null;
            dynamic @nam = null;
            dynamic @namlast = "";
            dynamic @newpath = "";
            dynamic @newpumptype = null;
            dynamic @obj = null;
            dynamic @objdoc = null;
            dynamic @objlayout = null;
            dynamic @objold = null;
            dynamic @objplot = null;
            dynamic @p = new double[3];
            dynamic @p1 = null;
            dynamic @p10 = null;
            dynamic @p11 = null;
            dynamic @p12 = null;
            dynamic @p13 = null;
            dynamic @p14 = null;
            dynamic @p15 = "";
            dynamic @p2 = null;
            dynamic @p3 = null;
            dynamic @p4 = null;
            dynamic @p5 = null;
            dynamic @p6 = null;
            dynamic @p7 = null;
            dynamic @p8 = null;
            dynamic @p9 = null;
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
            dynamic @yonghudingdanhao = null;
            dynamic @yonghumingcheng = "";
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
            // Normalized rule line 45
            // Normalized rule line 48
            @xlbook = Book;
            // Normalized rule line 49
            @xlsheet = Book.Worksheets["Sheet1"];
            // Normalized rule line 50
            @xlsheet1 = Book.Worksheets["PID"];
            // Normalized rule line 51
            @xlsheet2 = Book.Worksheets["input"];
            // Normalized rule line 53
            @i = Vb.Val(@xlsheet.Range[Vb.Cat("C", 19)].Value2);
            // Normalized rule line 54
            @savename = @xlsheet.Range[Vb.Cat("Q", @i)].Value2;
            // Normalized rule line 55
            @ver = Vb.Val(Vb.Left(@xlsheet2.Range[Vb.Cat("E", Vb.Sub(@i, 19))].Value2, 1));
            // Normalized rule line 56
            @shebeimingcheng = @xlsheet.Range[Vb.Cat("D", @i)].Value2;
            // Normalized rule line 57
            @yonghudingdanhao = @xlsheet.Range[Vb.Cat("C", 9)].Value2;
            // Normalized rule line 58
            @xiangmumingcheng = @xlsheet.Range[Vb.Cat("C", 8)].Value2;
            // Normalized rule line 59
            @yonghumingcheng = @xlsheet.Range[Vb.Cat("C", 7)].Value2;
            // Normalized rule line 60
            if (Vb.Bool(Vb.Ne(@xlsheet.Range[Vb.Cat("CN", @i)].Value2, ""))) {
                // Normalized rule line 61
                @newpumptype = Vb.Replace(@xlsheet.Range[Vb.Cat("C", @i)].Value2, Vb.Left(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 2), @xlsheet.Range[Vb.Cat("CN", @i)].Value2);
                // Normalized rule line 62
            } else {
                // Normalized rule line 63
                @newpumptype = @xlsheet.Range[Vb.Cat("C", @i)].Value2;
                // Normalized rule line 64
            }
            // Normalized rule line 65
            @j = Vb.Val(@xlsheet.Range[Vb.Cat("AU", @i)].Value2);
            // Normalized rule line 66
            @jiekou = @xlsheet.Range[Vb.Cat("BD", @i)].Value2;
            // Normalized rule line 67
            @pump = Vb.Left(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 2);
            // Normalized rule line 68
            if (Vb.Bool(Vb.And(Vb.Eq(@ver, 7), Vb.Eq(@pump, "ZE")))) {
                @pump = "ZE7";
            }
            // Normalized rule line 70
            @anchor = Vb.Right(@xlsheet.Range[Vb.Cat("AT", @i)].Value2, 7);
            // Normalized rule line 71
            @p1 = @xlsheet.Range[Vb.Cat("CH", @i)].Value2;
            // Normalized rule line 72
            @p2 = @xlsheet.Range[Vb.Cat("CI", @i)].Value2;
            // Normalized rule line 73
            @p3 = @xlsheet.Range[Vb.Cat("CJ", @i)].Value2;
            // Normalized rule line 74
            @p4 = @xlsheet.Range[Vb.Cat("CK", @i)].Value2;
            // Normalized rule line 75
            @p5 = @xlsheet.Range[Vb.Cat("CL", @i)].Value2;
            // Normalized rule line 76
            @p6 = @xlsheet.Range[Vb.Cat("CM", @i)].Value2;
            // Normalized rule line 77
            @lp1 = Vb.Left(@p1, 2);
            // Normalized rule line 78
            @lp2 = Vb.Left(@p2, 2);
            // Normalized rule line 79
            @lp3 = Vb.Left(@p3, 2);
            // Normalized rule line 80
            @lp4 = Vb.Left(@p4, 2);
            // Normalized rule line 81
            @lp5 = Vb.Left(@p5, 2);
            // Normalized rule line 82
            @lp6 = Vb.Left(@p6, 2);
            // Normalized rule line 84
            @p[Vb.Int(0)] = 0;
            // Normalized rule line 85
            @p[Vb.Int(1)] = 0;
            // Normalized rule line 86
            @p[Vb.Int(2)] = 0;
            // Normalized rule line 87
            @q[Vb.Int(0)] = 420;
            // Normalized rule line 88
            @q[Vb.Int(1)] = 297;
            // Normalized rule line 89
            @q[Vb.Int(2)] = 0;
            // Normalized rule line 90
            @r[Vb.Int(0)] = 361;
            // Normalized rule line 91
            @v[Vb.Int(0)] = 361;
            // Normalized rule line 92
            @w[Vb.Int(0)] = 361;
            // Normalized rule line 93
            @x[Vb.Int(0)] = 361;
            // Normalized rule line 94
            @r[Vb.Int(1)] = 63;
            // Normalized rule line 95
            @v[Vb.Int(1)] = 67;
            // Normalized rule line 96
            @x[Vb.Int(1)] = 71.5;
            // Normalized rule line 97
            @w[Vb.Int(1)] = 53.9;
            // Normalized rule line 98
            @r[Vb.Int(2)] = 0;
            // Normalized rule line 99
            @v[Vb.Int(2)] = 0;
            // Normalized rule line 100
            @w[Vb.Int(2)] = 0;
            // Normalized rule line 101
            @x[Vb.Int(2)] = 0;
            // Normalized rule line 102
            @t[Vb.Int(0)] = 0;
            // Normalized rule line 103
            @t[Vb.Int(1)] = Vb.Neg(4.5);
            // Normalized rule line 104
            @t[Vb.Int(2)] = 0;
            // Normalized rule line 105
            @u[Vb.Int(0)] = 3.5;
            // Normalized rule line 106
            @u[Vb.Int(1)] = 0;
            // Normalized rule line 107
            @u[Vb.Int(2)] = 0;
            // Normalized rule line 124
            @mtext1 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@r, 48, @shebeimingcheng);
            // Normalized rule line 127
            dynamic case1 = Vb.Lenb(@shebeimingcheng);
            // Normalized rule line 128
            if (Vb.Bool(Vb.Lt(case1, 18))) {
                // Normalized rule line 129
                @mtext1.ScaleEntity(@r, 1);
                // Normalized rule line 130
                @mtext1.Height = 2.3;
                // Normalized rule line 131
                @mtext1.Width = 48;
                // Normalized rule line 132
                @mtext1.Update();
                // Normalized rule line 133
            } else if (Vb.Bool(Vb.And(Vb.Ge(case1, 18), Vb.Le(case1, 36)))) {
                // Normalized rule line 134
                @mtext1.ScaleEntity(@r, 0.85);
                // Normalized rule line 135
                @mtext1.Height = 2;
                // Normalized rule line 136
                @mtext1.Width = 55;
                // Normalized rule line 137
                @mtext1.Update();
                // Normalized rule line 138
            } else if (Vb.Bool(Vb.Gt(case1, 36))) {
                // Normalized rule line 139
                @mtext1.ScaleEntity(@r, 0.6);
                // Normalized rule line 140
                @mtext1.Height = 1.65;
                // Normalized rule line 141
                @mtext1.Width = 75;
                // Normalized rule line 142
                @mtext1.Update();
                // Normalized rule line 143
            }
            // Normalized rule line 146
            @mtext2 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@v, 48, @xiangmumingcheng);
            // Normalized rule line 149
            dynamic case2 = Vb.Lenb(@xiangmumingcheng);
            // Normalized rule line 150
            if (Vb.Bool(Vb.Lt(case2, 18))) {
                // Normalized rule line 151
                @mtext2.ScaleEntity(@v, 1);
                // Normalized rule line 152
                @mtext2.Height = 2.3;
                // Normalized rule line 153
                @mtext2.Width = 48;
                // Normalized rule line 154
                @mtext2.Update();
                // Normalized rule line 155
            } else if (Vb.Bool(Vb.And(Vb.Ge(case2, 18), Vb.Le(case2, 36)))) {
                // Normalized rule line 156
                @mtext2.ScaleEntity(@v, 0.85);
                // Normalized rule line 157
                @mtext2.Height = 2;
                // Normalized rule line 158
                @mtext2.Width = 55;
                // Normalized rule line 159
                @mtext2.Update();
                // Normalized rule line 160
            } else if (Vb.Bool(Vb.Gt(case2, 36))) {
                // Normalized rule line 161
                @mtext2.ScaleEntity(@v, 0.6);
                // Normalized rule line 162
                @mtext2.Height = 1.65;
                // Normalized rule line 163
                @mtext2.Width = 75;
                // Normalized rule line 164
                @mtext2.Update();
                // Normalized rule line 165
            }
            // Normalized rule line 167
            @mtext3 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@w, 48, @yonghudingdanhao);
            // Normalized rule line 170
            dynamic case3 = Vb.Lenb(@yonghudingdanhao);
            // Normalized rule line 171
            if (Vb.Bool(Vb.Lt(case3, 18))) {
                // Normalized rule line 172
                @mtext3.ScaleEntity(@w, 1);
                // Normalized rule line 173
                @mtext3.Height = 2.3;
                // Normalized rule line 174
                @mtext3.Width = 48;
                // Normalized rule line 175
                @mtext3.Update();
                // Normalized rule line 176
            } else if (Vb.Bool(Vb.And(Vb.Ge(case3, 18), Vb.Le(case3, 36)))) {
                // Normalized rule line 177
                @mtext3.ScaleEntity(@w, 0.85);
                // Normalized rule line 178
                @mtext3.Height = 2;
                // Normalized rule line 179
                @mtext3.Width = 55;
                // Normalized rule line 180
                @mtext3.Update();
                // Normalized rule line 181
            } else if (Vb.Bool(Vb.Gt(case3, 36))) {
                // Normalized rule line 182
                @mtext3.ScaleEntity(@w, 0.6);
                // Normalized rule line 183
                @mtext3.Height = 1.6;
                // Normalized rule line 184
                @mtext3.Width = 70;
                // Normalized rule line 185
                @mtext3.Update();
                // Normalized rule line 186
            }
            // Normalized rule line 188
            @mtext4 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@x, 48, @yonghumingcheng);
            // Normalized rule line 191
            dynamic case4 = Vb.Lenb(@yonghumingcheng);
            // Normalized rule line 192
            if (Vb.Bool(Vb.Lt(case4, 18))) {
                // Normalized rule line 193
                @mtext4.ScaleEntity(@x, 1);
                // Normalized rule line 194
                @mtext4.Height = 2.3;
                // Normalized rule line 195
                @mtext4.Width = 48;
                // Normalized rule line 196
                @mtext4.Update();
                // Normalized rule line 197
            } else if (Vb.Bool(Vb.And(Vb.Ge(case4, 18), Vb.Le(case4, 36)))) {
                // Normalized rule line 198
                @mtext4.ScaleEntity(@x, 0.85);
                // Normalized rule line 199
                @mtext4.Height = 2;
                // Normalized rule line 200
                @mtext4.Width = 55;
                // Normalized rule line 201
                @mtext4.Update();
                // Normalized rule line 202
            } else if (Vb.Bool(Vb.Gt(case4, 36))) {
                // Normalized rule line 203
                @mtext4.ScaleEntity(@x, 0.6);
                // Normalized rule line 204
                @mtext4.Height = 1.6;
                // Normalized rule line 205
                @mtext4.Width = 70;
                // Normalized rule line 206
                @mtext4.Update();
                // Normalized rule line 207
            }
            // Normalized rule line 211
            @ptmin = @p;
            // Normalized rule line 212
            @ptmax = @q;
            // Normalized rule line 214
            if (Vb.Bool(Vb.Eq(@pump, "ZF"))) {
                // Normalized rule line 215
                @pump = "ZE";
                // Normalized rule line 216
            }
            // Normalized rule line 218
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@p1, "52"), Vb.Eq(@p2, "52")), Vb.Eq(@p3, "52")))) {
                // Normalized rule line 219
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "M"), Vb.Eq(@p3, "M")), Vb.Eq(@p4, "M")), Vb.Eq(@p5, "M")), Vb.Eq(@p6, "M")))) {
                    // Normalized rule line 220
                    @p7 = "52WM";
                    // Normalized rule line 221
                } else {
                    // Normalized rule line 222
                    @p7 = "52WOM";
                    // Normalized rule line 223
                }
                // Normalized rule line 224
            } else {
                // Normalized rule line 225
                @p7 = "";
                // Normalized rule line 226
            }
            // Normalized rule line 228
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@p1, "53A"), Vb.Eq(@p2, "53A")), Vb.Eq(@p3, "53A")))) {
                // Normalized rule line 229
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "M"), Vb.Eq(@p3, "M")), Vb.Eq(@p4, "M")), Vb.Eq(@p5, "M")), Vb.Eq(@p6, "M")))) {
                    // Normalized rule line 230
                    @p8 = "53AWM";
                    // Normalized rule line 231
                } else {
                    // Normalized rule line 232
                    @p8 = "53AWOM";
                    // Normalized rule line 233
                }
                // Normalized rule line 234
            } else {
                // Normalized rule line 235
                @p8 = "";
                // Normalized rule line 236
            }
            // Normalized rule line 238
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@p1, "53B"), Vb.Eq(@p2, "53B")), Vb.Eq(@p3, "53B")))) {
                // Normalized rule line 239
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "M"), Vb.Eq(@p3, "M")), Vb.Eq(@p4, "M")), Vb.Eq(@p5, "M")), Vb.Eq(@p6, "M")))) {
                    // Normalized rule line 240
                    @p14 = "53BWM";
                    // Normalized rule line 241
                } else {
                    // Normalized rule line 242
                    @p14 = "53BWOM";
                    // Normalized rule line 243
                }
                // Normalized rule line 244
            } else {
                // Normalized rule line 245
                @p14 = "";
                // Normalized rule line 246
            }
            // Normalized rule line 248
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "53C"), Vb.Eq(@p2, "53C")), Vb.Eq(@p3, "53C")), Vb.Eq(@p4, "53C")))) {
                // Normalized rule line 249
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "M"), Vb.Eq(@p3, "M")), Vb.Eq(@p4, "M")), Vb.Eq(@p5, "M")), Vb.Eq(@p6, "M")))) {
                    // Normalized rule line 250
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "11"), Vb.Eq(@p1, "21")), Vb.Eq(@p1, "23")), Vb.Eq(@p1, "32")), Vb.Eq(@p2, "62")), Vb.Eq(@p3, "62")))) {
                        // Normalized rule line 251
                        @p9 = "53CWM+";
                        // Normalized rule line 252
                    } else {
                        // Normalized rule line 253
                        @p9 = "53CSWM";
                        // Normalized rule line 254
                    }
                    // Normalized rule line 256
                } else {
                    // Normalized rule line 257
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "11"), Vb.Eq(@p1, "21")), Vb.Eq(@p1, "23")), Vb.Eq(@p1, "32")), Vb.Eq(@p2, "62")), Vb.Eq(@p3, "62")))) {
                        // Normalized rule line 258
                        @p9 = "53CWOM+";
                        // Normalized rule line 259
                    } else {
                        // Normalized rule line 260
                        @p9 = "53CSWOM";
                        // Normalized rule line 261
                    }
                    // Normalized rule line 262
                }
                // Normalized rule line 263
            } else {
                // Normalized rule line 264
                @p9 = "";
                // Normalized rule line 265
            }
            // Normalized rule line 271
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@p2, "C"), Vb.Eq(@p3, "C")), Vb.Eq(@p4, "C")))) {
                // Normalized rule line 272
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@p1, "21"), Vb.Eq(@p1, "11")), Vb.Eq(@p1, "23")))) {
                    // Normalized rule line 273
                    dynamic case5 = Vb.Ucase(@p1);
                    // Normalized rule line 274
                    if (Vb.Bool(Vb.Eq(case5, "21"))) {
                        // Normalized rule line 275
                        @p10 = "C+";
                        // Normalized rule line 276
                    } else if (Vb.Bool(Vb.Eq(case5, "11"))) {
                        // Normalized rule line 277
                        @p10 = "C+";
                        // Normalized rule line 278
                    } else if (Vb.Bool(Vb.Eq(case5, "23"))) {
                        // Normalized rule line 279
                        @p10 = "C++";
                        // Normalized rule line 280
                    }
                    // Normalized rule line 281
                } else {
                    // Normalized rule line 282
                    @p10 = "CS";
                    // Normalized rule line 283
                }
                // Normalized rule line 284
            } else {
                // Normalized rule line 285
                @p10 = "";
                // Normalized rule line 286
            }
            // Normalized rule line 292
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "S"), Vb.Eq(@p3, "S")), Vb.Eq(@p4, "S")), Vb.Eq(@p5, "S")), Vb.Eq(@p6, "S")))) {
                // Normalized rule line 293
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Eq(@p1, "11"), Vb.Eq(@p1, "21")), Vb.Eq(@p1, "23")))) {
                    // Normalized rule line 294
                    dynamic case6 = Vb.Ucase(@p1);
                    // Normalized rule line 295
                    if (Vb.Bool(Vb.Eq(case6, "21"))) {
                        // Normalized rule line 296
                        @p11 = "S+";
                        // Normalized rule line 297
                    } else if (Vb.Bool(Vb.Eq(case6, "11"))) {
                        // Normalized rule line 298
                        @p11 = "S+";
                        // Normalized rule line 299
                    } else if (Vb.Bool(Vb.Eq(case6, "23"))) {
                        // Normalized rule line 300
                        @p11 = "S++";
                        // Normalized rule line 301
                    }
                    // Normalized rule line 302
                } else {
                    // Normalized rule line 303
                    @p11 = "SS";
                    // Normalized rule line 304
                }
                // Normalized rule line 305
            } else {
                // Normalized rule line 306
                @p11 = "";
                // Normalized rule line 307
            }
            // Normalized rule line 310
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "A"), Vb.Eq(@p3, "A")), Vb.Eq(@p4, "A")), Vb.Eq(@p5, "A")))) {
                // Normalized rule line 311
                @p12 = "A";
                // Normalized rule line 312
            } else {
                // Normalized rule line 313
                @p12 = "";
                // Normalized rule line 314
            }
            // Normalized rule line 316
            if (Vb.Bool(Vb.And(Vb.And(Vb.And(Vb.And(Vb.Ne(@p2, "M"), Vb.Ne(@p3, "M")), Vb.Ne(@p4, "M")), Vb.Ne(@p5, "M")), Vb.Ne(@p6, "M")))) {
                // Normalized rule line 317
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p10, "C+"), Vb.Eq(@p10, "C++")), Vb.Eq(@p10, "CS")), Vb.Eq(@p11, "S+")), Vb.Eq(@p11, "SS")), Vb.Eq(@p11, "S++")), Vb.Eq(@p12, "A")))) {
                    // Normalized rule line 318
                    @p13 = "M";
                    // Normalized rule line 319
                } else {
                    // Normalized rule line 320
                    @p13 = "";
                    // Normalized rule line 321
                }
                // Normalized rule line 322
            }
            // Normalized rule line 324
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "72"), Vb.Eq(@p2, "72")), Vb.Eq(@p3, "72")), Vb.Eq(@p4, "72")), Vb.Eq(@p5, "72")), Vb.Eq(@p6, "72")))) {
                // Normalized rule line 325
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "76"), Vb.Eq(@p3, "76")), Vb.Eq(@p4, "76")), Vb.Eq(@p5, "76")), Vb.Eq(@p6, "76")))) {
                    // Normalized rule line 326
                    @p15 = "72DS";
                    // Normalized rule line 327
                } else {
                    // Normalized rule line 328
                    @p15 = "72US";
                    // Normalized rule line 329
                }
                // Normalized rule line 330
            } else {
                // Normalized rule line 331
                @p15 = "";
                // Normalized rule line 332
            }
            // Normalized rule line 334
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "32"), Vb.Eq(@p2, "32")), Vb.Eq(@p3, "32")), Vb.Eq(@p1, "62")), Vb.Eq(@p2, "62")), Vb.Eq(@p3, "62")), Vb.Eq(@p4, "62")))) {
                // Normalized rule line 335
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.And(Vb.Eq(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 2), "AC"), Vb.Ne(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 3), "3AC")), Vb.And(Vb.Eq(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 2), "AS"), Vb.Ne(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 3), "3AS"))), Vb.Eq(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 2), "CS")))) {
                    // Normalized rule line 336
                    Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = true;
                    // Normalized rule line 337
                    Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = false;
                    // Normalized rule line 338
                } else {
                    // Normalized rule line 339
                    Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = false;
                    // Normalized rule line 340
                    Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = true;
                    // Normalized rule line 341
                }
                // Normalized rule line 342
            }
            // Normalized rule line 347
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "M"), Vb.Eq(@p3, "M")), Vb.Eq(@p4, "M")), Vb.Eq(@p5, "M")), Vb.Eq(@p6, "M")))) {
                // Normalized rule line 348
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "21"), Vb.Eq(@p1, "23")), Vb.Eq(@p1, "41")), Vb.Eq(@p1, "32")), Vb.Eq(@p2, "32")), Vb.Eq(@p1, "62")), Vb.Eq(@p2, "62")), Vb.Eq(@p3, "62")))) {
                    // Normalized rule line 349
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "A"), Vb.Eq(@p3, "A")), Vb.Eq(@p4, "A")), Vb.Eq(@p5, "A")), Vb.Eq(@p2, "C")), Vb.Eq(@p3, "C")), Vb.Eq(@p4, "C")), Vb.Eq(@p5, "C")), Vb.Eq(@p6, "C")), Vb.Eq(@p2, "S")), Vb.Eq(@p3, "S")), Vb.Eq(@p4, "S")), Vb.Eq(@p5, "S")), Vb.Eq(@p6, "S")))) {
                        // Normalized rule line 350
                        Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = true;
                        // Normalized rule line 351
                        Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = false;
                        // Normalized rule line 352
                    } else {
                        // Normalized rule line 353
                        Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = false;
                        // Normalized rule line 354
                        Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = true;
                        // Normalized rule line 355
                    }
                    // Normalized rule line 356
                }
                // Normalized rule line 357
            }
            // Normalized rule line 359
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "52"), Vb.Eq(@p2, "52")), Vb.Eq(@p3, "52")), Vb.Eq(@p4, "52")), Vb.Eq(@p5, "52")), Vb.Eq(@lp1, "53")), Vb.Eq(@lp2, "53")), Vb.Eq(@lp3, "53")), Vb.Eq(@lp4, "53")), Vb.Eq(@lp5, "53")), Vb.Eq(@p1, "54")), Vb.Eq(@p2, "54")), Vb.Eq(@p3, "54")), Vb.Eq(@p1, "72")), Vb.Eq(@p2, "72")), Vb.Eq(@p3, "72")), Vb.Eq(@p4, "72")), Vb.Eq(@p1, "74")), Vb.Eq(@p2, "74")), Vb.Eq(@p3, "74")), Vb.Eq(@p4, "74")), Vb.Eq(@p2, "75")), Vb.Eq(@p3, "75")), Vb.Eq(@p4, "75")), Vb.Eq(@p5, "75")), Vb.Eq(@p2, "76")), Vb.Eq(@p3, "76")), Vb.Eq(@p4, "76")), Vb.Eq(@p5, "76")))) {
                // Normalized rule line 360
                if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p2, "M"), Vb.Eq(@p3, "M")), Vb.Eq(@p4, "M")), Vb.Eq(@p5, "M")), Vb.Eq(@p6, "M")))) {
                    // Normalized rule line 361
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "21"), Vb.Eq(@p1, "23")), Vb.Eq(@p1, "41")), Vb.Eq(@p2, "A")), Vb.Eq(@p3, "A")), Vb.Eq(@p4, "A")), Vb.Eq(@p5, "A")), Vb.Eq(@p2, "C")), Vb.Eq(@p3, "C")), Vb.Eq(@p4, "C")), Vb.Eq(@p5, "C")), Vb.Eq(@p6, "C")), Vb.Eq(@p2, "S")), Vb.Eq(@p3, "S")), Vb.Eq(@p4, "S")), Vb.Eq(@p5, "S")), Vb.Eq(@p6, "S")))) {
                        // Normalized rule line 362
                        Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = true;
                        // Normalized rule line 363
                        Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = false;
                        // Normalized rule line 365
                    } else {
                        // Normalized rule line 366
                        Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = false;
                        // Normalized rule line 367
                        Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = true;
                        // Normalized rule line 368
                    }
                    // Normalized rule line 369
                } else {
                    // Normalized rule line 370
                    if (Vb.Bool(Vb.Or(Vb.Or(Vb.And(Vb.Eq(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 2), "AC"), Vb.Ne(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 3), "3AC")), Vb.And(Vb.Eq(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 2), "AS"), Vb.Ne(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 3), "3AS"))), Vb.Eq(Vb.Right(Vb.Replace(@xlsheet.Range[Vb.Cat("CG", @i)].Value2, "+", ""), 2), "CS")))) {
                        // Normalized rule line 371
                        Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = true;
                        // Normalized rule line 372
                        Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = false;
                        // Normalized rule line 374
                    } else {
                        // Normalized rule line 375
                        Drawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = false;
                        // Normalized rule line 376
                        Drawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = true;
                        // Normalized rule line 377
                    }
                    // Normalized rule line 378
                }
                // Normalized rule line 379
            }
            // Normalized rule line 382
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Ne(@p7, ""), Vb.Ne(@p8, "")), Vb.Eq(@p14, "53BWM")))) {
                // Normalized rule line 383
                Drawing.Application.ActiveDocument.Layers.Item("AO1").LayerOn = true;
                // Normalized rule line 384
            } else {
                // Normalized rule line 385
                Drawing.Application.ActiveDocument.Layers.Item("AO1").LayerOn = false;
                // Normalized rule line 386
            }
            // Normalized rule line 388
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p9, "53CWM+"), Vb.Eq(@p9, "53CWOM+")), Vb.Eq(@p14, "53BWOM")), Vb.Eq(@p7, "52WM")), Vb.Eq(@p8, "53AWM")), Vb.Eq(@p14, "53BWM")))) {
                // Normalized rule line 389
                Drawing.Application.ActiveDocument.Layers.Item("AO2").LayerOn = true;
                // Normalized rule line 390
            } else {
                // Normalized rule line 391
                Drawing.Application.ActiveDocument.Layers.Item("AO2").LayerOn = false;
                // Normalized rule line 392
            }
            // Normalized rule line 394
            if (Vb.Bool(Vb.Eq(@p9, "53CWM+"))) {
                // Normalized rule line 395
                Drawing.Application.ActiveDocument.Layers.Item("AO3").LayerOn = true;
                // Normalized rule line 396
            } else {
                // Normalized rule line 397
                Drawing.Application.ActiveDocument.Layers.Item("AO3").LayerOn = false;
                // Normalized rule line 398
            }
            // Normalized rule line 400
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Ne(@p7, ""), Vb.Ne(@p8, "")), Vb.Eq(@p14, "53BWM")))) {
                // Normalized rule line 401
                Drawing.Application.ActiveDocument.Layers.Item("AF1").LayerOn = false;
                // Normalized rule line 402
            } else {
                // Normalized rule line 403
                Drawing.Application.ActiveDocument.Layers.Item("AF1").LayerOn = true;
                // Normalized rule line 404
            }
            // Normalized rule line 406
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p9, "53CWM+"), Vb.Eq(@p9, "53CWOM+")), Vb.Eq(@p14, "53BWOM")), Vb.Eq(@p7, "52WM")), Vb.Eq(@p8, "53AWM")), Vb.Eq(@p14, "53BWM")))) {
                // Normalized rule line 407
                Drawing.Application.ActiveDocument.Layers.Item("AF2").LayerOn = false;
                // Normalized rule line 408
            } else {
                // Normalized rule line 409
                Drawing.Application.ActiveDocument.Layers.Item("AF2").LayerOn = true;
                // Normalized rule line 410
            }
            // Normalized rule line 412
            if (Vb.Bool(Vb.Eq(@p9, "53CWM+"))) {
                // Normalized rule line 413
                Drawing.Application.ActiveDocument.Layers.Item("AF3").LayerOn = false;
                // Normalized rule line 414
            } else {
                // Normalized rule line 415
                Drawing.Application.ActiveDocument.Layers.Item("AF3").LayerOn = true;
                // Normalized rule line 416
            }
            // Normalized rule line 418
            if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 18)].Value2), 1))) {
                // Normalized rule line 419
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "PD", 1, 1, 1, 0);
                // Normalized rule line 420
            }
            // Normalized rule line 424
            if (Vb.Bool(Vb.Or(Vb.Or(Vb.Or(Vb.Or(Vb.Eq(@p1, "21"), Vb.Eq(@p1, "23")), Vb.Eq(@p1, "41")), Vb.Eq(@p2, "41")), Vb.Ne(@p11, "")))) {
                // Normalized rule line 425
                if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 18)].Value2), 1))) {
                    // Normalized rule line 426
                    Drawing.Application.ActiveDocument.Layers.Item("PDO").LayerOn = true;
                    // Normalized rule line 427
                    Drawing.Application.ActiveDocument.Layers.Item("PDF").LayerOn = false;
                    // Normalized rule line 428
                } else {
                    // Normalized rule line 429
                    Drawing.Application.ActiveDocument.Layers.Item("PDF").LayerOn = true;
                    // Normalized rule line 430
                    Drawing.Application.ActiveDocument.Layers.Item("PDO").LayerOn = false;
                    // Normalized rule line 431
                }
                // Normalized rule line 432
            }
            // Normalized rule line 436
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @pump, 1, 1, 1, 0);
            // Normalized rule line 437
            if (Vb.Bool(Vb.And(Vb.Eq(@p1, "32"), Vb.Eq(@p2, "53B")))) {
                // Normalized rule line 438
                InsertPidPlan(@u, @p1);
                // Normalized rule line 439
            } else {
                // Normalized rule line 440
                InsertPidPlan(@p, @p1);
                // Normalized rule line 441
            }
            // Normalized rule line 442
            InsertPidPlan(@p, @p2);
            // Normalized rule line 443
            InsertPidPlan(@p, @p3);
            // Normalized rule line 444
            InsertPidPlan(@p, @p4);
            // Normalized rule line 445
            InsertPidPlan(@p, @p5);
            // Normalized rule line 446
            InsertPidPlan(@p, @p6);
            // Normalized rule line 447
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p7, 1, 1, 1, 0);
            // Normalized rule line 448
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p8, 1, 1, 1, 0);
            // Normalized rule line 449
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p9, 1, 1, 1, 0);
            // Normalized rule line 450
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p10, 1, 1, 1, 0);
            // Normalized rule line 451
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p11, 1, 1, 1, 0);
            // Normalized rule line 452
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p12, 1, 1, 1, 0);
            // Normalized rule line 453
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p13, 1, 1, 1, 0);
            // Normalized rule line 454
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p14, 1, 1, 1, 0);
            // Normalized rule line 455
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @p15, 1, 1, 1, 0);
            // Normalized rule line 459
            if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 18)].Value2), 1))) {
                // Normalized rule line 461
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "jibenjiekou", 1, 1, 1, 0);
                // Normalized rule line 462
                @t[Vb.Int(1)] = 0;
                // Normalized rule line 463
            } else {
                // Normalized rule line 464
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "jibenjiekouwod", 1, 1, 1, 0);
                // Normalized rule line 466
            }
            // Normalized rule line 467
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@t, @jiekou, 1, 1, 1, 0);
            // Normalized rule line 470
            if (Vb.Bool(Vb.And(Vb.Ne(@xlsheet1.Range[Vb.Cat("G", Vb.Sub(@i, 30))].Value2, ""), Vb.Ne(@xlsheet1.Range[Vb.Cat("H", Vb.Sub(@i, 30))].Value2, "")))) {
                // Normalized rule line 471
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "BVBD", 1, 1, 1, 0);
                // Normalized rule line 472
            }
            // Normalized rule line 477
            @count = Drawing.ActiveLayout.Block.Count;
            // Normalized rule line 478
            for (@index = 0; Vb.Le(@index, Vb.Sub(@count, 1)); @index = Vb.Add(@index, 1)) {
                // Normalized rule line 479
                @tempstring = Drawing.ActiveLayout.Block.Item(@index).ObjectName;
                // Normalized rule line 480
                if (Vb.Bool(Vb.Eq(@tempstring, "AcDbBlockReference"))) {
                    // Normalized rule line 481
                    if (Vb.Bool(true)) {
                        // Normalized rule line 482
                        @blockattributes = Drawing.ActiveLayout.Block.Item(@index).GetAttributes();
                        // Normalized rule line 483
                        foreach (dynamic each7 in @blockattributes) {
                            @attrib = each7;
                            // Normalized rule line 484
                            @tagname = @attrib.TagString;
                            // Normalized rule line 485
                            dynamic case8 = Vb.Ucase(@tagname);
                            // Normalized rule line 496
                            if (Vb.Bool(Vb.Eq(case8, "SHEBEIWEIHAO"))) {
                                // Normalized rule line 498
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("B", @i)].Value2);
                                // Normalized rule line 502
                            } else if (Vb.Bool(Vb.Eq(case8, "DINGDANHAO_SULZER"))) {
                                // Normalized rule line 504
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 10)].Value2);
                                // Normalized rule line 505
                            } else if (Vb.Bool(Vb.Eq(case8, "GUANLUFANGAN"))) {
                                // Normalized rule line 507
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("E", @i)].Value2);
                                // Normalized rule line 510
                            } else if (Vb.Bool(Vb.Eq(case8, "TYPE_PUMP"))) {
                                // Normalized rule line 512
                                @attrib.TextString = Vb.Text(Vb.Cat("                       ", @newpumptype));
                                // Normalized rule line 513
                            } else if (Vb.Bool(Vb.Eq(case8, "DATE"))) {
                                // Normalized rule line 515
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 6)].Text);
                                // Normalized rule line 516
                            } else if (Vb.Bool(Vb.Eq(case8, "DESIGN"))) {
                                // Normalized rule line 518
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 4)].Value2);
                                // Normalized rule line 519
                            } else if (Vb.Bool(Vb.Eq(case8, "CHECK"))) {
                                // Normalized rule line 521
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 5)].Value2);
                                // Normalized rule line 523
                            } else if (Vb.Bool(Vb.Eq(case8, "NAME"))) {
                                // Normalized rule line 525
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 4)].Value2);
                                // Normalized rule line 528
                            } else if (Vb.Bool(Vb.Eq(case8, "SCI"))) {
                                // Normalized rule line 530
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AL", @i)].Value2);
                                // Normalized rule line 531
                            } else if (Vb.Bool(Vb.Eq(case8, "SCO"))) {
                                // Normalized rule line 533
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AL", @i)].Value2);
                                // Normalized rule line 534
                            } else if (Vb.Bool(Vb.Eq(case8, "FCI"))) {
                                // Normalized rule line 536
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AO", @i)].Value2);
                                // Normalized rule line 537
                            } else if (Vb.Bool(Vb.Eq(case8, "FCO"))) {
                                // Normalized rule line 539
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AO", @i)].Value2);
                                // Normalized rule line 540
                            } else if (Vb.Bool(Vb.Eq(case8, "SQI"))) {
                                // Normalized rule line 542
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AP", @i)].Value2);
                                // Normalized rule line 543
                            } else if (Vb.Bool(Vb.Eq(case8, "SQO"))) {
                                // Normalized rule line 545
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AP", @i)].Value2);
                                // Normalized rule line 546
                            } else if (Vb.Bool(Vb.Eq(case8, "FQI"))) {
                                // Normalized rule line 548
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AQ", @i)].Value2);
                                // Normalized rule line 549
                            } else if (Vb.Bool(Vb.Eq(case8, "FQO"))) {
                                // Normalized rule line 551
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AQ", @i)].Value2);
                                // Normalized rule line 552
                            } else if (Vb.Bool(Vb.Eq(case8, "SFI"))) {
                                // Normalized rule line 554
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AR", @i)].Value2);
                                // Normalized rule line 555
                            } else if (Vb.Bool(Vb.Eq(case8, "FFI"))) {
                                // Normalized rule line 557
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AS", @i)].Value2);
                                // Normalized rule line 558
                            } else if (Vb.Bool(Vb.Eq(case8, "SBI"))) {
                                // Normalized rule line 560
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AV", @i)].Value2);
                                // Normalized rule line 561
                            } else if (Vb.Bool(Vb.Eq(case8, "SBO"))) {
                                // Normalized rule line 563
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AV", @i)].Value2);
                                // Normalized rule line 564
                            } else if (Vb.Bool(Vb.Eq(case8, "FBI"))) {
                                // Normalized rule line 566
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AW", @i)].Value2);
                                // Normalized rule line 567
                            } else if (Vb.Bool(Vb.Eq(case8, "FBO"))) {
                                // Normalized rule line 569
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AW", @i)].Value2);
                                // Normalized rule line 570
                            } else if (Vb.Bool(Vb.Eq(case8, "SD"))) {
                                // Normalized rule line 572
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("S", @i)].Value2);
                                // Normalized rule line 573
                            } else if (Vb.Bool(Vb.Eq(case8, "SS"))) {
                                // Normalized rule line 575
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("T", @i)].Value2);
                                // Normalized rule line 576
                            } else if (Vb.Bool(Vb.Eq(case8, "FSD"))) {
                                // Normalized rule line 578
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("U", @i)].Value2);
                                // Normalized rule line 579
                            } else if (Vb.Bool(Vb.Eq(case8, "SPD"))) {
                                // Normalized rule line 581
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("V", @i)].Value2);
                                // Normalized rule line 582
                            } else if (Vb.Bool(Vb.Eq(case8, "FPD"))) {
                                // Normalized rule line 584
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("W", @i)].Value2);
                                // Normalized rule line 585
                            } else if (Vb.Bool(Vb.Eq(case8, "DFX"))) {
                                // Normalized rule line 587
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AX", @i)].Value2);
                                // Normalized rule line 588
                            } else if (Vb.Bool(Vb.Eq(case8, "DFY"))) {
                                // Normalized rule line 590
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AY", @i)].Value2);
                                // Normalized rule line 591
                            } else if (Vb.Bool(Vb.Eq(case8, "DFZ"))) {
                                // Normalized rule line 593
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AZ", @i)].Value2);
                                // Normalized rule line 594
                            } else if (Vb.Bool(Vb.Eq(case8, "SFX"))) {
                                // Normalized rule line 596
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BA", @i)].Value2);
                                // Normalized rule line 597
                            } else if (Vb.Bool(Vb.Eq(case8, "SFY"))) {
                                // Normalized rule line 599
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BB", @i)].Value2);
                                // Normalized rule line 600
                            } else if (Vb.Bool(Vb.Eq(case8, "SFZ"))) {
                                // Normalized rule line 602
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BC", @i)].Value2);
                                // Normalized rule line 603
                            } else if (Vb.Bool(Vb.Eq(case8, "DMX"))) {
                                // Normalized rule line 605
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BK", @i)].Value2);
                                // Normalized rule line 606
                            } else if (Vb.Bool(Vb.Eq(case8, "DMY"))) {
                                // Normalized rule line 608
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BL", @i)].Value2);
                                // Normalized rule line 609
                            } else if (Vb.Bool(Vb.Eq(case8, "DMZ"))) {
                                // Normalized rule line 611
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BM", @i)].Value2);
                                // Normalized rule line 612
                            } else if (Vb.Bool(Vb.Eq(case8, "SMX"))) {
                                // Normalized rule line 614
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BN", @i)].Value2);
                                // Normalized rule line 615
                            } else if (Vb.Bool(Vb.Eq(case8, "SMY"))) {
                                // Normalized rule line 617
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BO", @i)].Value2);
                                // Normalized rule line 618
                            } else if (Vb.Bool(Vb.Eq(case8, "SMZ"))) {
                                // Normalized rule line 620
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BP", @i)].Value2);
                                // Normalized rule line 621
                            } else if (Vb.Bool(Vb.Eq(case8, "SAD2"))) {
                                // Normalized rule line 623
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BJ", @i)].Value2);
                                // Normalized rule line 624
                            } else if (Vb.Bool(Vb.Eq(case8, "FAD2"))) {
                                // Normalized rule line 626
                                @attrib.TextString = Vb.Text("NPTF");
                                // Normalized rule line 628
                            } else if (Vb.Bool(Vb.Eq(case8, "SBV"))) {
                                // Normalized rule line 630
                                @attrib.TextString = Vb.Text(@xlsheet1.Range[Vb.Cat("G", Vb.Sub(@i, 30))].Value2);
                                // Normalized rule line 631
                            } else if (Vb.Bool(Vb.Eq(case8, "SBD"))) {
                                // Normalized rule line 633
                                @attrib.TextString = Vb.Text(@xlsheet1.Range[Vb.Cat("H", Vb.Sub(@i, 30))].Value2);
                                // Normalized rule line 634
                            } else if (Vb.Bool(Vb.Eq(case8, "FBV"))) {
                                // Normalized rule line 636
                                @attrib.TextString = Vb.Text(@xlsheet1.Range[Vb.Cat("I", Vb.Sub(@i, 30))].Value2);
                                // Normalized rule line 637
                            } else if (Vb.Bool(Vb.Eq(case8, "FBD"))) {
                                // Normalized rule line 639
                                @attrib.TextString = Vb.Text(@xlsheet1.Range[Vb.Cat("I", Vb.Sub(@i, 30))].Value2);
                                // Normalized rule line 645
                            }
                            // Normalized rule line 646
                        }
                        // Normalized rule line 647
                    }
                    // Normalized rule line 648
                }
                // Normalized rule line 649
            }
        }
    }
}
