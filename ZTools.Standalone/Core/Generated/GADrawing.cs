// Ported from migration/legacy; see tools/port_vba.py and migration/README.md.
using System;
namespace ZTools.Engine {
    internal sealed partial class SubmissionEngine {
        public void BuildGA() {
            dynamic @anchor = "";
            dynamic @attrib = null;
            dynamic @blockattributes = null;
            dynamic @bold = false;
            
            dynamic @charset = 0;
            dynamic @cname1 = null;
            dynamic @cname2 = null;
            dynamic @cname3 = null;
            dynamic @cool = "";
            dynamic @count = null;
            dynamic @counter = 0;
            dynamic @cvalue1 = null;
            dynamic @cvalue2 = null;
            dynamic @cvalue3 = "";
            dynamic @drain = "";
            dynamic @drainza = "";
            dynamic @exchanger = null;
            dynamic @gongsimingcheng = null;
            dynamic @i = null;
            dynamic @i1 = 0;
            dynamic @index = null;
            dynamic @italic = false;
            dynamic @j = null;
            dynamic @jiekou = null;
            dynamic @k = null;
            dynamic @last = "";
            dynamic @m = 0;
            dynamic @mscale = null;
            dynamic @mtext1 = null;
            dynamic @mtext2 = null;
            dynamic @mtext3 = null;
            dynamic @mtext4 = null;
            dynamic @myobj = null;
            dynamic @myset = null;
            dynamic @nam = null;
            dynamic @newpath = "";
            dynamic @newpumptype = "";
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
            dynamic @ss = null;
            dynamic @stz = null;
            dynamic @t = new double[3];
            dynamic @tagname = null;
            dynamic @tempstring = null;
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
            // Normalized rule line 47
            @xlbook = Book;
            // Normalized rule line 48
            @xlsheet = Book.Worksheets["Sheet1"];
            // Normalized rule line 49
            @xlsheet1 = Book.Worksheets["input"];
            // Normalized rule line 51
            @i = Vb.Val(@xlsheet.Range[Vb.Cat("C", 19)].Value2);
            // Normalized rule line 52
            @savename = @xlsheet.Range[Vb.Cat("P", @i)].Value2;
            // Normalized rule line 53
            @shebeimingcheng = @xlsheet.Range[Vb.Cat("D", @i)].Value2;
            // Normalized rule line 54
            @yonghudingdanhao = @xlsheet.Range[Vb.Cat("C", 9)].Value2;
            // Normalized rule line 55
            @xiangmumingcheng = @xlsheet.Range[Vb.Cat("C", 8)].Value2;
            // Normalized rule line 56
            @yonghumingcheng = @xlsheet.Range[Vb.Cat("C", 7)].Value2;
            // Normalized rule line 57
            if (Vb.Bool(Vb.Ne(@xlsheet.Range[Vb.Cat("CN", @i)].Value2, ""))) {
                // Normalized rule line 58
                @newpumptype = Vb.Replace(@xlsheet.Range[Vb.Cat("C", @i)].Value2, Vb.Left(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 2), @xlsheet.Range[Vb.Cat("CN", @i)].Value2);
                // Normalized rule line 59
            } else {
                // Normalized rule line 60
                @newpumptype = @xlsheet.Range[Vb.Cat("C", @i)].Value2;
                // Normalized rule line 61
            }
            // Normalized rule line 62
            @j = Vb.Val(@xlsheet.Range[Vb.Cat("AU", @i)].Value2);
            // Normalized rule line 63
            @jiekou = @xlsheet.Range[Vb.Cat("BD", @i)].Value2;
            // Normalized rule line 64
            @pump = Vb.Left(@xlsheet.Range[Vb.Cat("C", @i)].Value2, 2);
            // Normalized rule line 65
            @anchor = Vb.Right(@xlsheet.Range[Vb.Cat("AT", @i)].Value2, 7);
            // Normalized rule line 66
            @exchanger = @xlsheet.Range[Vb.Cat("BE", @i)].Value2;
            // Normalized rule line 67
            @reservoir = @xlsheet.Range[Vb.Cat("BF", @i)].Value2;
            // Normalized rule line 68
            @side = @xlsheet.Range[Vb.Cat("BI", @i)].Value2;
            // Normalized rule line 69
            @cool = @xlsheet.Range[Vb.Cat("AL", @i)].Value2;
            // Normalized rule line 70
            @gongsimingcheng = @xlsheet1.Range[Vb.Cat("K", 10)].Value2;
            // Normalized rule line 71
            @ver = Vb.Val(Vb.Left(@xlsheet1.Range[Vb.Cat("E", Vb.Sub(@i, 19))].Value2, 1));
            // Normalized rule line 72
            @p[Vb.Int(0)] = 0;
            // Normalized rule line 73
            @p[Vb.Int(1)] = 0;
            // Normalized rule line 74
            @p[Vb.Int(2)] = 0;
            // Normalized rule line 75
            @q[Vb.Int(0)] = 420;
            // Normalized rule line 76
            @q[Vb.Int(1)] = 297;
            // Normalized rule line 77
            @q[Vb.Int(2)] = 0;
            // Normalized rule line 78
            @r[Vb.Int(0)] = 361;
            // Normalized rule line 79
            @v[Vb.Int(0)] = 361;
            // Normalized rule line 80
            @w[Vb.Int(0)] = 361;
            // Normalized rule line 81
            @x[Vb.Int(0)] = 361;
            // Normalized rule line 82
            @r[Vb.Int(1)] = 63;
            // Normalized rule line 83
            @v[Vb.Int(1)] = 67;
            // Normalized rule line 84
            @x[Vb.Int(1)] = 71.5;
            // Normalized rule line 85
            @w[Vb.Int(1)] = 53.9;
            // Normalized rule line 86
            @r[Vb.Int(2)] = 0;
            // Normalized rule line 87
            @v[Vb.Int(2)] = 0;
            // Normalized rule line 88
            @w[Vb.Int(2)] = 0;
            // Normalized rule line 89
            @x[Vb.Int(2)] = 0;
            // Normalized rule line 90
            @t[Vb.Int(0)] = 0;
            // Normalized rule line 91
            @t[Vb.Int(1)] = Vb.Neg(4.5);
            // Normalized rule line 92
            @t[Vb.Int(2)] = 0;
            // Normalized rule line 97
            @ptmin = @p;
            // Normalized rule line 98
            @ptmax = @q;
            // Normalized rule line 100
            if (Vb.Bool(Vb.Or(Vb.Eq(@pump, "ZA"), Vb.Eq(@pump, "CZ")))) {
                // Normalized rule line 101
                @k = 7;
                // Normalized rule line 102
                @m = 7;
                // Normalized rule line 103
                @s[Vb.Int(0)] = 12.6394;
                // Normalized rule line 104
                @s[Vb.Int(1)] = Vb.Neg(4.4978);
                // Normalized rule line 105
                @s[Vb.Int(2)] = 0;
                // Normalized rule line 106
            }
            // Normalized rule line 107
            if (Vb.Bool(Vb.Eq(@pump, "ZE"))) {
                // Normalized rule line 108
                if (Vb.Bool(Vb.Eq(@ver, 7))) {
                    @pump = "ZE7";
                }
                // Normalized rule line 109
                @k = 8;
                // Normalized rule line 110
                @m = 8;
                // Normalized rule line 111
                @s[Vb.Int(0)] = 0;
                // Normalized rule line 112
                @s[Vb.Int(1)] = 0;
                // Normalized rule line 113
                @s[Vb.Int(2)] = 0;
                // Normalized rule line 114
            }
            // Normalized rule line 115
            if (Vb.Bool(Vb.Eq(@pump, "ZF"))) {
                // Normalized rule line 116
                @m = 10;
                // Normalized rule line 117
                @k = 8;
                // Normalized rule line 118
                @s[Vb.Int(0)] = 0;
                // Normalized rule line 119
                @s[Vb.Int(1)] = 0;
                // Normalized rule line 120
                @s[Vb.Int(2)] = 0;
                // Normalized rule line 121
            }
            // Normalized rule line 138
            for (@i1 = 4; Vb.Le(@i1, 18); @i1 = Vb.Add(@i1, 1)) {
                // Normalized rule line 145
                ChangeField(@xlsheet.Range[Vb.Cat("G", @i1)].Value2, @xlsheet.Range[Vb.Cat("H", @i1)].Value2);
                // Normalized rule line 146
                ChangeField(@xlsheet.Range[Vb.Cat("I", @i1)].Value2, @xlsheet.Range[Vb.Cat("J", @i1)].Value2);
                // Normalized rule line 147
                ChangeField(@xlsheet.Range[Vb.Cat("K", @i1)].Value2, @xlsheet.Range[Vb.Cat("L", @i1)].Value2);
                // Normalized rule line 149
            }
            // Normalized rule line 156
            @mtext1 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@r, 48, @shebeimingcheng);
            // Normalized rule line 159
            dynamic case1 = Vb.Lenb(@shebeimingcheng);
            // Normalized rule line 160
            if (Vb.Bool(Vb.Lt(case1, 18))) {
                // Normalized rule line 161
                @mtext1.ScaleEntity(@r, 1);
                // Normalized rule line 162
                @mtext1.Height = 2.3;
                // Normalized rule line 163
                @mtext1.Width = 48;
                // Normalized rule line 164
                @mtext1.Update();
                // Normalized rule line 165
            } else if (Vb.Bool(Vb.And(Vb.Ge(case1, 18), Vb.Le(case1, 36)))) {
                // Normalized rule line 166
                @mtext1.ScaleEntity(@r, 0.85);
                // Normalized rule line 167
                @mtext1.Height = 2;
                // Normalized rule line 168
                @mtext1.Width = 55;
                // Normalized rule line 169
                @mtext1.Update();
                // Normalized rule line 170
            } else if (Vb.Bool(Vb.Gt(case1, 36))) {
                // Normalized rule line 171
                @mtext1.ScaleEntity(@r, 0.6);
                // Normalized rule line 172
                @mtext1.Height = 1.65;
                // Normalized rule line 173
                @mtext1.Width = 75;
                // Normalized rule line 174
                @mtext1.Update();
                // Normalized rule line 175
            }
            // Normalized rule line 178
            @mtext2 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@v, 48, @xiangmumingcheng);
            // Normalized rule line 181
            dynamic case2 = Vb.Lenb(@xiangmumingcheng);
            // Normalized rule line 182
            if (Vb.Bool(Vb.Lt(case2, 18))) {
                // Normalized rule line 183
                @mtext2.ScaleEntity(@v, 1);
                // Normalized rule line 184
                @mtext2.Height = 2.3;
                // Normalized rule line 185
                @mtext2.Width = 48;
                // Normalized rule line 186
                @mtext2.Update();
                // Normalized rule line 187
            } else if (Vb.Bool(Vb.And(Vb.Ge(case2, 18), Vb.Le(case2, 36)))) {
                // Normalized rule line 188
                @mtext2.ScaleEntity(@v, 0.85);
                // Normalized rule line 189
                @mtext2.Height = 2;
                // Normalized rule line 190
                @mtext2.Width = 55;
                // Normalized rule line 191
                @mtext2.Update();
                // Normalized rule line 192
            } else if (Vb.Bool(Vb.And(Vb.Ge(case2, 37), Vb.Le(case2, 54)))) {
                // Normalized rule line 193
                @mtext2.ScaleEntity(@v, 0.6);
                // Normalized rule line 194
                @mtext2.Height = 1.65;
                // Normalized rule line 195
                @mtext2.Width = 75;
                // Normalized rule line 196
                @mtext2.Update();
                // Normalized rule line 197
            } else if (Vb.Bool(Vb.And(Vb.Ge(case2, 55), Vb.Le(case2, 60)))) {
                // Normalized rule line 198
                @mtext2.ScaleEntity(@v, 0.5);
                // Normalized rule line 199
                @mtext2.Height = 1;
                // Normalized rule line 200
                @mtext2.Width = 75;
                // Normalized rule line 201
                @mtext2.Update();
                // Normalized rule line 202
            } else if (Vb.Bool(Vb.Gt(case2, 60))) {
                // Normalized rule line 203
                @mtext2.ScaleEntity(@v, 0.4);
                // Normalized rule line 204
                @mtext2.Height = 0.8;
                // Normalized rule line 205
                @mtext2.Width = 65;
                // Normalized rule line 206
                @mtext2.Update();
                // Normalized rule line 207
            }
            // Normalized rule line 209
            @mtext3 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@w, 48, @yonghudingdanhao);
            // Normalized rule line 212
            dynamic case3 = Vb.Lenb(@yonghudingdanhao);
            // Normalized rule line 213
            if (Vb.Bool(Vb.Lt(case3, 18))) {
                // Normalized rule line 214
                @mtext3.ScaleEntity(@w, 1);
                // Normalized rule line 215
                @mtext3.Height = 2.3;
                // Normalized rule line 216
                @mtext3.Width = 48;
                // Normalized rule line 217
                @mtext3.Update();
                // Normalized rule line 218
            } else if (Vb.Bool(Vb.And(Vb.Ge(case3, 18), Vb.Le(case3, 36)))) {
                // Normalized rule line 219
                @mtext3.ScaleEntity(@w, 0.85);
                // Normalized rule line 220
                @mtext3.Height = 2;
                // Normalized rule line 221
                @mtext3.Width = 55;
                // Normalized rule line 222
                @mtext3.Update();
                // Normalized rule line 223
            } else if (Vb.Bool(Vb.Gt(case3, 36))) {
                // Normalized rule line 224
                @mtext3.ScaleEntity(@w, 0.6);
                // Normalized rule line 225
                @mtext3.Height = 1.6;
                // Normalized rule line 226
                @mtext3.Width = 70;
                // Normalized rule line 227
                @mtext3.Update();
                // Normalized rule line 228
            }
            // Normalized rule line 230
            @mtext4 = Drawing.Application.ActiveDocument.ModelSpace.AddMText(@x, 48, @yonghumingcheng);
            // Normalized rule line 233
            dynamic case4 = Vb.Lenb(@yonghumingcheng);
            // Normalized rule line 234
            if (Vb.Bool(Vb.Lt(case4, 18))) {
                // Normalized rule line 235
                @mtext4.ScaleEntity(@x, 1);
                // Normalized rule line 236
                @mtext4.Height = 2.3;
                // Normalized rule line 237
                @mtext4.Width = 48;
                // Normalized rule line 238
                @mtext4.Update();
                // Normalized rule line 239
            } else if (Vb.Bool(Vb.And(Vb.Ge(case4, 18), Vb.Le(case4, 36)))) {
                // Normalized rule line 240
                @mtext4.ScaleEntity(@x, 0.85);
                // Normalized rule line 241
                @mtext4.Height = 2;
                // Normalized rule line 242
                @mtext4.Width = 55;
                // Normalized rule line 243
                @mtext4.Update();
                // Normalized rule line 244
            } else if (Vb.Bool(Vb.Gt(case4, 36))) {
                // Normalized rule line 245
                @mtext4.ScaleEntity(@x, 0.6);
                // Normalized rule line 246
                @mtext4.Height = 1.6;
                // Normalized rule line 247
                @mtext4.Width = 70;
                // Normalized rule line 248
                @mtext4.Update();
                // Normalized rule line 249
            }
            // Normalized rule line 254
            if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 18)].Value2), 1))) {
                // Normalized rule line 255
                @drain = @xlsheet.Range[Vb.Cat("BR", @i)].Value2;
                // Normalized rule line 256
                if (Vb.Bool(Vb.Ne(@pump, "CZ"))) {
                    // Normalized rule line 257
                    Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("Drain", @m), 1, 1, 1, 0);
                    // Normalized rule line 258
                }
                // Normalized rule line 259
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @drain, 1, 1, 1, 0);
                // Normalized rule line 260
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "jibenjiekou", 1, 1, 1, 0);
                // Normalized rule line 261
                @t[Vb.Int(1)] = 0;
                // Normalized rule line 262
            } else {
                // Normalized rule line 263
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "jibenjiekouwod", 1, 1, 1, 0);
                // Normalized rule line 265
            }
            // Normalized rule line 267
            if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 16)].Value2), 1))) {
                // Normalized rule line 268
                Drawing.Application.ActiveDocument.Layers.Item("matingflange").LayerOn = true;
                // Normalized rule line 269
            } else {
                // Normalized rule line 270
                Drawing.Application.ActiveDocument.Layers.Item("matingflange").LayerOn = false;
                // Normalized rule line 271
            }
            // Normalized rule line 274
            if (Vb.Bool(Vb.Eq(@exchanger, "Y"))) {
                // Normalized rule line 275
                if (Vb.Bool(Vb.Ne(@side, "L"))) {
                    // Normalized rule line 276
                    Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat(Vb.Cat("huanreqi", @m), "you"), 1, 1, 1, 0);
                    // Normalized rule line 277
                } else {
                    // Normalized rule line 278
                    if (Vb.Bool(Vb.Eq(@reservoir, "Y"))) {
                        // Normalized rule line 279
                        Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat(Vb.Cat("huanreqi", @m), "you"), 1, 1, 1, 0);
                        // Normalized rule line 280
                    } else {
                        // Normalized rule line 281
                        Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat(Vb.Cat("huanreqi", @k), "zuo"), 1, 1, 1, 0);
                        // Normalized rule line 282
                    }
                    // Normalized rule line 283
                }
                // Normalized rule line 284
            }
            // Normalized rule line 286
            if (Vb.Bool(Vb.Eq(Vb.Val(@xlsheet.Range[Vb.Cat("C", 17)].Value2), 1))) {
                // Normalized rule line 287
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("youwurunhua", @m), 1, 1, 1, 0);
                // Normalized rule line 288
            }
            // Normalized rule line 291
            if (Vb.Bool(Vb.Eq(@reservoir, "Y"))) {
                // Normalized rule line 292
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("hongxiguan", @m), 1, 1, 1, 0);
                // Normalized rule line 293
            }
            // Normalized rule line 295
            if (Vb.Bool(Vb.Ne(@cool, "NO"))) {
                // Normalized rule line 296
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("lengqueshui", @k), 1, 1, 1, 0);
                // Normalized rule line 297
            }
            // Normalized rule line 299
            if (Vb.Bool(Vb.Ne(@xlsheet.Range[Vb.Cat("AR", @i)].Value2, "NO"))) {
                // Normalized rule line 300
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("tuxingFI", @k), 1, 1, 1, 0);
                // Normalized rule line 301
            }
            // Normalized rule line 303
            if (Vb.Bool(Vb.Ne(@xlsheet.Range[Vb.Cat("AP", @i)].Value2, "NO"))) {
                // Normalized rule line 304
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("tuxingQ", @k), 1, 1, 1, 0);
                // Normalized rule line 305
            }
            // Normalized rule line 307
            if (Vb.Bool(Vb.Ne(@xlsheet.Range[Vb.Cat("AV", @i)].Value2, "NO"))) {
                // Normalized rule line 308
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("tuxingB", @k), 1, 1, 1, 0);
                // Normalized rule line 309
            }
            // Normalized rule line 311
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @pump, 1, 1, 1, 0);
            // Normalized rule line 312
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, Vb.Cat("jibenchicun", @k), 1, 1, 1, 0);
            // Normalized rule line 313
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@s, Vb.Cat("geduan", @j), 1, 1, 1, 0);
            // Normalized rule line 314
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, "nozzle", 1, 1, 1, 0);
            // Normalized rule line 315
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@t, @jiekou, 1, 1, 1, 0);
            // Normalized rule line 316
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(@p, @anchor, 1, 1, 1, 0);
            // Normalized rule line 320
            @count = Drawing.ActiveLayout.Block.Count;
            // Normalized rule line 321
            for (@index = 0; Vb.Le(@index, Vb.Sub(@count, 1)); @index = Vb.Add(@index, 1)) {
                // Normalized rule line 322
                @tempstring = Drawing.ActiveLayout.Block.Item(@index).ObjectName;
                // Normalized rule line 323
                if (Vb.Bool(Vb.Eq(@tempstring, "AcDbBlockReference"))) {
                    // Normalized rule line 324
                    if (Vb.Bool(true)) {
                        // Normalized rule line 325
                        @blockattributes = Drawing.ActiveLayout.Block.Item(@index).GetAttributes();
                        // Normalized rule line 326
                        foreach (dynamic each5 in @blockattributes) {
                            @attrib = each5;
                            // Normalized rule line 327
                            @tagname = @attrib.TagString;
                            // Normalized rule line 328
                            dynamic case6 = Vb.Ucase(@tagname);
                            // Normalized rule line 339
                            if (Vb.Bool(Vb.Eq(case6, "SHEBEIWEIHAO"))) {
                                // Normalized rule line 341
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("B", @i)].Value2);
                                // Normalized rule line 345
                            } else if (Vb.Bool(Vb.Eq(case6, "DINGDANHAO_SULZER"))) {
                                // Normalized rule line 347
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 10)].Value2);
                                // Normalized rule line 348
                            } else if (Vb.Bool(Vb.Eq(case6, "GUANLUFANGAN"))) {
                                // Normalized rule line 350
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("E", @i)].Value2);
                                // Normalized rule line 353
                            } else if (Vb.Bool(Vb.Eq(case6, "TYPE_PUMP"))) {
                                // Normalized rule line 355
                                @attrib.TextString = Vb.Text(Vb.Cat(" ", @newpumptype));
                                // Normalized rule line 356
                            } else if (Vb.Bool(Vb.Eq(case6, "DATE"))) {
                                // Normalized rule line 358
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 6)].Text);
                                // Normalized rule line 359
                            } else if (Vb.Bool(Vb.Eq(case6, "DESIGN"))) {
                                // Normalized rule line 361
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 4)].Value2);
                                // Normalized rule line 362
                            } else if (Vb.Bool(Vb.Eq(case6, "CHECK"))) {
                                // Normalized rule line 364
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 5)].Value2);
                                // Normalized rule line 365
                            } else if (Vb.Bool(Vb.Eq(case6, "PUMP"))) {
                                // Normalized rule line 367
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("F", @i)].Value2));
                                // Normalized rule line 368
                            } else if (Vb.Bool(Vb.Eq(case6, "MOTOR"))) {
                                // Normalized rule line 370
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("G", @i)].Value2));
                                // Normalized rule line 371
                            } else if (Vb.Bool(Vb.Eq(case6, "BASE"))) {
                                // Normalized rule line 373
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("H", @i)].Value2));
                                // Normalized rule line 374
                            } else if (Vb.Bool(Vb.Eq(case6, "ACCESSORY"))) {
                                // Normalized rule line 376
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("I", @i)].Value2));
                                // Normalized rule line 377
                            } else if (Vb.Bool(Vb.Eq(case6, "TOTAL"))) {
                                // Normalized rule line 379
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("J", @i)].Value2));
                                // Normalized rule line 380
                            } else if (Vb.Bool(Vb.Eq(case6, "DIANJIXINGHAO"))) {
                                // Normalized rule line 382
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("K", @i)].Value2);
                                // Normalized rule line 383
                            } else if (Vb.Bool(Vb.Eq(case6, "DIANYA"))) {
                                // Normalized rule line 385
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("L", @i)].Value2);
                                // Normalized rule line 386
                            } else if (Vb.Bool(Vb.Eq(case6, "EDINGGONGLV"))) {
                                // Normalized rule line 388
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("M", @i)].Value2));
                                // Normalized rule line 389
                            } else if (Vb.Bool(Vb.Eq(case6, "DIANJICHANGJIA"))) {
                                // Normalized rule line 391
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("N", @i)].Value2);
                                // Normalized rule line 393
                            } else if (Vb.Bool(Vb.Eq(case6, "ZHUANSU"))) {
                                // Normalized rule line 395
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("O", @i)].Value2));
                                // Normalized rule line 396
                            } else if (Vb.Bool(Vb.Eq(case6, "COMPANY"))) {
                                // Normalized rule line 398
                                @attrib.TextString = Vb.Text(@gongsimingcheng);
                                // Normalized rule line 400
                            } else if (Vb.Bool(Vb.Eq(case6, "SCI"))) {
                                // Normalized rule line 402
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AL", @i)].Value2);
                                // Normalized rule line 403
                            } else if (Vb.Bool(Vb.Eq(case6, "SCO"))) {
                                // Normalized rule line 405
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AL", @i)].Value2);
                                // Normalized rule line 406
                            } else if (Vb.Bool(Vb.Eq(case6, "FCI"))) {
                                // Normalized rule line 408
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AO", @i)].Value2);
                                // Normalized rule line 409
                            } else if (Vb.Bool(Vb.Eq(case6, "FCO"))) {
                                // Normalized rule line 411
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AO", @i)].Value2);
                                // Normalized rule line 412
                            } else if (Vb.Bool(Vb.Eq(case6, "SQI"))) {
                                // Normalized rule line 414
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AP", @i)].Value2);
                                // Normalized rule line 415
                            } else if (Vb.Bool(Vb.Eq(case6, "SQO"))) {
                                // Normalized rule line 417
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AP", @i)].Value2);
                                // Normalized rule line 418
                            } else if (Vb.Bool(Vb.Eq(case6, "FQI"))) {
                                // Normalized rule line 420
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AQ", @i)].Value2);
                                // Normalized rule line 421
                            } else if (Vb.Bool(Vb.Eq(case6, "FQO"))) {
                                // Normalized rule line 423
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AQ", @i)].Value2);
                                // Normalized rule line 424
                            } else if (Vb.Bool(Vb.Eq(case6, "SFI"))) {
                                // Normalized rule line 426
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AR", @i)].Value2);
                                // Normalized rule line 427
                            } else if (Vb.Bool(Vb.Eq(case6, "FFI"))) {
                                // Normalized rule line 429
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AS", @i)].Value2);
                                // Normalized rule line 430
                            } else if (Vb.Bool(Vb.Eq(case6, "SBI"))) {
                                // Normalized rule line 432
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AV", @i)].Value2);
                                // Normalized rule line 433
                            } else if (Vb.Bool(Vb.Eq(case6, "SBO"))) {
                                // Normalized rule line 435
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AV", @i)].Value2);
                                // Normalized rule line 436
                            } else if (Vb.Bool(Vb.Eq(case6, "FBI"))) {
                                // Normalized rule line 438
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AW", @i)].Value2);
                                // Normalized rule line 439
                            } else if (Vb.Bool(Vb.Eq(case6, "FBO"))) {
                                // Normalized rule line 441
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AW", @i)].Value2);
                                // Normalized rule line 442
                            } else if (Vb.Bool(Vb.Eq(case6, "SD"))) {
                                // Normalized rule line 444
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("S", @i)].Value2);
                                // Normalized rule line 445
                            } else if (Vb.Bool(Vb.Eq(case6, "SS"))) {
                                // Normalized rule line 447
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("T", @i)].Value2);
                                // Normalized rule line 448
                            } else if (Vb.Bool(Vb.Eq(case6, "FSD"))) {
                                // Normalized rule line 450
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("U", @i)].Value2);
                                // Normalized rule line 451
                            } else if (Vb.Bool(Vb.Eq(case6, "SPD"))) {
                                // Normalized rule line 453
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("V", @i)].Value2);
                                // Normalized rule line 454
                            } else if (Vb.Bool(Vb.Eq(case6, "FPD"))) {
                                // Normalized rule line 456
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("W", @i)].Value2);
                                // Normalized rule line 457
                            } else if (Vb.Bool(Vb.Eq(case6, "DFX"))) {
                                // Normalized rule line 459
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AX", @i)].Value2);
                                // Normalized rule line 460
                            } else if (Vb.Bool(Vb.Eq(case6, "DFY"))) {
                                // Normalized rule line 462
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AY", @i)].Value2);
                                // Normalized rule line 463
                            } else if (Vb.Bool(Vb.Eq(case6, "DFZ"))) {
                                // Normalized rule line 465
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AZ", @i)].Value2);
                                // Normalized rule line 466
                            } else if (Vb.Bool(Vb.Eq(case6, "SFX"))) {
                                // Normalized rule line 468
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BA", @i)].Value2);
                                // Normalized rule line 469
                            } else if (Vb.Bool(Vb.Eq(case6, "SFY"))) {
                                // Normalized rule line 471
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BB", @i)].Value2);
                                // Normalized rule line 472
                            } else if (Vb.Bool(Vb.Eq(case6, "SFZ"))) {
                                // Normalized rule line 474
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BC", @i)].Value2);
                                // Normalized rule line 475
                            } else if (Vb.Bool(Vb.Eq(case6, "DMX"))) {
                                // Normalized rule line 477
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BK", @i)].Value2);
                                // Normalized rule line 478
                            } else if (Vb.Bool(Vb.Eq(case6, "DMY"))) {
                                // Normalized rule line 480
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BL", @i)].Value2);
                                // Normalized rule line 481
                            } else if (Vb.Bool(Vb.Eq(case6, "DMZ"))) {
                                // Normalized rule line 483
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BM", @i)].Value2);
                                // Normalized rule line 484
                            } else if (Vb.Bool(Vb.Eq(case6, "SMX"))) {
                                // Normalized rule line 486
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BN", @i)].Value2);
                                // Normalized rule line 487
                            } else if (Vb.Bool(Vb.Eq(case6, "SMY"))) {
                                // Normalized rule line 489
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BO", @i)].Value2);
                                // Normalized rule line 490
                            } else if (Vb.Bool(Vb.Eq(case6, "SMZ"))) {
                                // Normalized rule line 492
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BP", @i)].Value2);
                                // Normalized rule line 493
                            } else if (Vb.Bool(Vb.Eq(case6, "SAD2"))) {
                                // Normalized rule line 495
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BJ", @i)].Value2);
                                // Normalized rule line 496
                            } else if (Vb.Bool(Vb.Eq(case6, "FAD2"))) {
                                // Normalized rule line 498
                                @attrib.TextString = Vb.Text("NPTF");
                                // Normalized rule line 500
                            } else if (Vb.Bool(Vb.Eq(case6, "LF"))) {
                                // Normalized rule line 502
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BT", @i)].Value2);
                                // Normalized rule line 503
                            } else if (Vb.Bool(Vb.Eq(case6, "LB"))) {
                                // Normalized rule line 505
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BV", @i)].Value2);
                                // Normalized rule line 506
                            } else if (Vb.Bool(Vb.Eq(case6, "LQ"))) {
                                // Normalized rule line 508
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BX", @i)].Value2);
                                // Normalized rule line 509
                            } else if (Vb.Bool(Vb.Eq(case6, "HF"))) {
                                // Normalized rule line 511
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BS", @i)].Value2);
                                // Normalized rule line 512
                            } else if (Vb.Bool(Vb.Eq(case6, "HB"))) {
                                // Normalized rule line 514
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BU", @i)].Value2);
                                // Normalized rule line 515
                            } else if (Vb.Bool(Vb.Eq(case6, "HQ"))) {
                                // Normalized rule line 517
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BW", @i)].Value2);
                                // Normalized rule line 518
                            } else if (Vb.Bool(Vb.Eq(case6, "LEX"))) {
                                // Normalized rule line 520
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BG", @i)].Value2);
                                // Normalized rule line 521
                            } else if (Vb.Bool(Vb.Eq(case6, "LRE"))) {
                                // Normalized rule line 523
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("BH", @i)].Value2);
                                // Normalized rule line 524
                            } else if (Vb.Bool(Vb.Eq(case6, "NAME"))) {
                                // Normalized rule line 526
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("C", 4)].Value2);
                                // Normalized rule line 528
                            } else if (Vb.Bool(Vb.Eq(case6, "E"))) {
                                // Normalized rule line 530
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AJ", @i)].Value2));
                                // Normalized rule line 532
                            } else if (Vb.Bool(Vb.Eq(case6, "F"))) {
                                // Normalized rule line 534
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AK", @i)].Value2));
                                // Normalized rule line 535
                            } else if (Vb.Bool(Vb.Eq(case6, "Y"))) {
                                // Normalized rule line 537
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AM", @i)].Value2));
                                // Normalized rule line 538
                            } else if (Vb.Bool(Vb.Eq(case6, "K"))) {
                                // Normalized rule line 540
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AC", @i)].Value2));
                                // Normalized rule line 541
                            } else if (Vb.Bool(Vb.Eq(case6, "H_DRAIN"))) {
                                // Normalized rule line 543
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AN", @i)].Value2));
                                // Normalized rule line 544
                            } else if (Vb.Bool(Vb.Eq(case6, "H2"))) {
                                // Normalized rule line 546
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AI", @i)].Value2));
                                // Normalized rule line 547
                            } else if (Vb.Bool(Vb.Eq(case6, "L3"))) {
                                // Normalized rule line 549
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AB", @i)].Value2));
                                // Normalized rule line 550
                            } else if (Vb.Bool(Vb.Eq(case6, "L4"))) {
                                // Normalized rule line 552
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("BQ", @i)].Value2));
                                // Normalized rule line 553
                            } else if (Vb.Bool(Vb.Eq(case6, "L"))) {
                                // Normalized rule line 555
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("Y", @i)].Value2));
                                // Normalized rule line 556
                            } else if (Vb.Bool(Vb.Eq(case6, "B1"))) {
                                // Normalized rule line 558
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AE", @i)].Value2));
                                // Normalized rule line 559
                            } else if (Vb.Bool(Vb.Eq(case6, "B"))) {
                                // Normalized rule line 561
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AD", @i)].Value2));
                                // Normalized rule line 562
                            } else if (Vb.Bool(Vb.Eq(case6, "I"))) {
                                // Normalized rule line 564
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AF", @i)].Value2));
                                // Normalized rule line 565
                            } else if (Vb.Bool(Vb.Eq(case6, "H_BASE"))) {
                                // Normalized rule line 567
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AG", @i)].Value2));
                                // Normalized rule line 568
                            } else if (Vb.Bool(Vb.Eq(case6, "H_PUMP"))) {
                                // Normalized rule line 570
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("X", @i)].Value2));
                                // Normalized rule line 571
                            } else if (Vb.Bool(Vb.Eq(case6, "H_ANCHOR"))) {
                                // Normalized rule line 573
                                @attrib.TextString = Vb.Text(Vb.Right(@xlsheet.Range[Vb.Cat("AT", @i)].Value2, 3));
                                // Normalized rule line 574
                            } else if (Vb.Bool(Vb.Eq(case6, "BOLTS"))) {
                                // Normalized rule line 576
                                @attrib.TextString = Vb.Text(@xlsheet.Range[Vb.Cat("AT", @i)].Value2);
                                // Normalized rule line 578
                            } else if (Vb.Bool(Vb.Eq(case6, "L2"))) {
                                // Normalized rule line 580
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("AA", @i)].Value2));
                                // Normalized rule line 581
                            } else if (Vb.Bool(Vb.Eq(case6, "L1"))) {
                                // Normalized rule line 583
                                @attrib.TextString = Vb.Text(Vb.Val(@xlsheet.Range[Vb.Cat("Z", @i)].Value2));
                                // Normalized rule line 585
                            }
                            // Normalized rule line 586
                        }
                        // Normalized rule line 587
                    }
                    // Normalized rule line 588
                }
                // Normalized rule line 589
            }
        }
    }
}
