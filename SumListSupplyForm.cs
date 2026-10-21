using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SESDTools
{
    public partial class SumListSupplyForm : Form
    {
        private Boolean textboxHasText1 = false;
        private Boolean textboxHasText2 = false;
        private Boolean textboxHasText3 = false;
        private Boolean textboxHasText4 = false;
        public SumListSupplyForm()
        {
            InitializeComponent();
            Box1_Enter(new object(), new System.EventArgs());
            Box1_Leave(new object(), new System.EventArgs());
            Box2_Enter(new object(), new System.EventArgs());
            Box2_Leave(new object(), new System.EventArgs());
            Box3_Enter(new object(), new System.EventArgs());
            Box3_Leave(new object(), new System.EventArgs());
            Box4_Enter(new object(), new System.EventArgs());
            Box4_Leave(new object(), new System.EventArgs());
            textboxHasText1 = false;
            textboxHasText2 = false;
            textboxHasText3 = false;
            textboxHasText4 = false;
        }

        private void reset()
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox8.Text = "";
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            radioButton1.Checked = true;
            radioButton3.Checked = true;
            radioButton5.Checked = true;
            radioButton7.Checked = true;
            radioButton9.Checked = true;
            radioButton11.Checked = true;
            radioButton13.Checked = true;
            radioButton15.Checked = true;
            radioButton17.Checked = true;
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;

            label1.BackColor = Color.Transparent;
            label2.BackColor = Color.Transparent;
            label3.BackColor = Color.Transparent;
            label4.BackColor = Color.Transparent;
            label5.BackColor = Color.Transparent;
            label6.BackColor = Color.Transparent;
            label7.BackColor = Color.Transparent;
            label8.BackColor = Color.Transparent;
            label9.BackColor = Color.Transparent;
            label10.BackColor = Color.Transparent;
            label11.BackColor = Color.Transparent;
            label12.BackColor = Color.Transparent;
            label13.BackColor = Color.Transparent;
            label14.BackColor = Color.Transparent;
            label15.BackColor = Color.Transparent;
            label16.BackColor = Color.Transparent;
            label17.BackColor = Color.Transparent;
            label18.BackColor = Color.Transparent;
            label19.BackColor = Color.Transparent;
            label20.BackColor = Color.Transparent;
            label21.BackColor = Color.Transparent;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Dispose(true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Dispose(true);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            reset();
            Box1_Leave(new object(), new System.EventArgs());
            Box2_Leave(new object(), new System.EventArgs());
            Box3_Leave(new object(), new System.EventArgs());
            Box4_Leave(new object(), new System.EventArgs());
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //textbox获得焦点
        private void Box1_Enter(object sender, EventArgs e)
        {
            if (textboxHasText1 == false)
            {
                textBox1.Text = "";
            }
            textBox1.ForeColor = Color.Black;
        }
        //textbox失去焦点
        private void Box1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "请手动输入";
                textBox1.ForeColor = Color.LightGray;
                textboxHasText1 = false;
            }
            else
            {
                textboxHasText1 = true;
            }
        }


        //textbox获得焦点
        private void Box2_Enter(object sender, EventArgs e)
        {
            if (textboxHasText2 == false)
            {
                textBox2.Text = "";
            }
            textBox2.ForeColor = Color.Black;
        }
        //textbox失去焦点
        private void Box2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.Text = "请手动输入";
                textBox2.ForeColor = Color.LightGray;
                textboxHasText2 = false;
            }
            else
            {
                textboxHasText2 = true;
            }
        }

        //textbox获得焦点
        private void Box3_Enter(object sender, EventArgs e)
        {
            if (textboxHasText3 == false)
            {
                textBox3.Text = "";
            }
            textBox3.ForeColor = Color.Black;
        }
        //textbox失去焦点
        private void Box3_Leave(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                textBox3.Text = "请手动输入";
                textBox3.ForeColor = Color.LightGray;
                textboxHasText3 = false;
            }
            else
            {
                textboxHasText3 = true;
            }
        }

        //textbox获得焦点
        private void Box4_Enter(object sender, EventArgs e)
        {
            if (textboxHasText4 == false)
            {
                textBox4.Text = "";
            }
            textBox4.ForeColor = Color.Black;
        }
        //textbox失去焦点
        private void Box4_Leave(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                textBox4.Text = "请手动输入";
                textBox4.ForeColor = Color.LightGray;
                textboxHasText4 = false;
            }
            else
            {
                textboxHasText4 = true;
            }
        }



        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }


        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton16_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton17_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton18_CheckedChanged(object sender, EventArgs e)
        {

        }



        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            reset();
            Box1_Enter(new object(), new System.EventArgs());
            Box1_Leave(new object(), new System.EventArgs());
            Box2_Enter(new object(), new System.EventArgs());
            Box2_Leave(new object(), new System.EventArgs());
            Box3_Enter(new object(), new System.EventArgs());
            Box3_Leave(new object(), new System.EventArgs());
            Box4_Enter(new object(), new System.EventArgs());
            Box4_Leave(new object(), new System.EventArgs());
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {

        }


        private void txt_GotFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box1_Enter(new object(), new System.EventArgs());
        }

        private void txt_LostFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box1_Leave(new object(), new System.EventArgs());
        }

        private void txt2_GotFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box2_Enter(new object(), new System.EventArgs());
        }

        private void txt2_LostFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box2_Leave(new object(), new System.EventArgs());
        }

        private void txt3_GotFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box3_Enter(new object(), new System.EventArgs());
        }

        private void txt3_LostFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box3_Leave(new object(), new System.EventArgs());
        }

        private void txt4_GotFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box4_Enter(new object(), new System.EventArgs());
        }

        private void txt4_LostFocus(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            Box4_Leave(new object(), new System.EventArgs());
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                label1.BackColor = Color.Silver;
                label2.BackColor = Color.Silver;
                label3.BackColor = Color.Silver;
                label4.BackColor = Color.Silver;
                label5.BackColor = Color.Silver;
            }
            else
            {
                if (checkBox2.Checked != true)
                {
                    label2.BackColor = Color.Transparent;
                    label3.BackColor = Color.Transparent;
                    label4.BackColor = Color.Transparent;
                    label5.BackColor = Color.Transparent;
                }
                label1.BackColor = Color.Transparent;

            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                label6.BackColor = Color.Silver;
                label2.BackColor = Color.Silver;
                label3.BackColor = Color.Silver;
                label4.BackColor = Color.Silver;
                label5.BackColor = Color.Silver;
            }
            else
            {
                if (checkBox1.Checked != true)
                {
                    label2.BackColor = Color.Transparent;
                }
                if (checkBox1.Checked != true && checkBox3.Checked != true)
                {
                    label3.BackColor = Color.Transparent;
                    label4.BackColor = Color.Transparent;
                    label5.BackColor = Color.Transparent;
                }
                if (checkBox3.Checked != true)
                {
                    label6.BackColor = Color.Transparent;
                }

            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                label6.BackColor = Color.Silver;
                label7.BackColor = Color.Silver;
                label3.BackColor = Color.Silver;
                label4.BackColor = Color.Silver;
                label5.BackColor = Color.Silver;
            }
            else
            {
                if (checkBox2.Checked != true)
                {
                    label3.BackColor = Color.Transparent;
                }
                if (checkBox2.Checked != true && checkBox4.Checked != true)
                {
                    label4.BackColor = Color.Transparent;
                    label5.BackColor = Color.Transparent;
                    label6.BackColor = Color.Transparent;
                }
                if (checkBox4.Checked != true)
                {
                    label7.BackColor = Color.Transparent;
                }
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                label6.BackColor = Color.Silver;
                label7.BackColor = Color.Silver;
                label8.BackColor = Color.Silver;
                label4.BackColor = Color.Silver;
                label5.BackColor = Color.Silver;
            }
            else
            {
                if (checkBox3.Checked != true)
                {
                    label4.BackColor = Color.Transparent;
                }
                if (checkBox3.Checked != true && checkBox5.Checked != true)
                {
                    label5.BackColor = Color.Transparent;
                    label6.BackColor = Color.Transparent;
                    label7.BackColor = Color.Transparent;
                }
                if (checkBox5.Checked != true)
                {
                    label8.BackColor = Color.Transparent;
                }
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                label6.BackColor = Color.Silver;
                label7.BackColor = Color.Silver;
                label8.BackColor = Color.Silver;
                label9.BackColor = Color.Silver;
                label5.BackColor = Color.Silver;
            }
            else
            {
                if (checkBox4.Checked != true)
                {
                    label5.BackColor = Color.Transparent;
                }
                if (checkBox4.Checked != true && checkBox6.Checked != true)
                {
                    label6.BackColor = Color.Transparent;
                    label7.BackColor = Color.Transparent;
                    label8.BackColor = Color.Transparent;
                }
                if (checkBox6.Checked != true)
                {
                    label9.BackColor = Color.Transparent;
                }
            }

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                label6.BackColor = Color.Silver;
                label7.BackColor = Color.Silver;
                label8.BackColor = Color.Silver;
                label9.BackColor = Color.Silver;
                label10.BackColor = Color.Silver;
            }
            else
            {
                if (checkBox5.Checked != true)
                {

                }
                if (checkBox5.Checked != true && checkBox7.Checked != true)
                {
                    label6.BackColor = Color.Transparent;
                    label7.BackColor = Color.Transparent;
                    label8.BackColor = Color.Transparent;
                    label9.BackColor = Color.Transparent;
                    label10.BackColor = Color.Transparent;
                }
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                label6.BackColor = Color.Silver;
                label7.BackColor = Color.Silver;
                label8.BackColor = Color.Silver;
                label9.BackColor = Color.Silver;
                label10.BackColor = Color.Silver;
                label11.BackColor = Color.Silver;
                label12.BackColor = Color.Silver;
                label13.BackColor = Color.Silver;
                label14.BackColor = Color.Silver;
                label15.BackColor = Color.Silver;
                label16.BackColor = Color.Silver;
                label17.BackColor = Color.Silver;
                label18.BackColor = Color.Silver;
                label19.BackColor = Color.Silver;
                label20.BackColor = Color.Silver;
                label21.BackColor = Color.Silver;
             }
            else
            {
                if (checkBox6.Checked != true)
                {
                    label6.BackColor = Color.Transparent;
                    label7.BackColor = Color.Transparent;
                    label8.BackColor = Color.Transparent;
                    label9.BackColor = Color.Transparent;
                    label10.BackColor = Color.Transparent;
                }
                label11.BackColor = Color.Transparent;
                label12.BackColor = Color.Transparent;
                label13.BackColor = Color.Transparent;
                label14.BackColor = Color.Transparent;
                label15.BackColor = Color.Transparent;
                label16.BackColor = Color.Transparent;
                label17.BackColor = Color.Transparent;
                label18.BackColor = Color.Transparent;
                label19.BackColor = Color.Transparent;
                label20.BackColor = Color.Transparent;
                label21.BackColor = Color.Transparent;
            }
        }
    }
}
