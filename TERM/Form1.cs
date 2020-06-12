using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using exportWord = Microsoft.Office.Interop.Word;

namespace TERM
{
    public partial class Form1 : MaterialForm
    {
        public Form1()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.Grey800, Accent.Red200, TextShade.WHITE);

        }

        Decimal zotvet, zotvet2, mgotvet, mgotvet2, alotvet, alotvet2, al_2_otvet, al_2_otvet2, fesotvet, fesotvet2, notvet, notvet2; //ответы к задачам
        Decimal otvet, otvet2, acotvet, acotvet2, cotvet, cotvet2, ac_2_otvet, ac_2_otvet2, potvet, potvet2, ootvet, ootvet2, hotvet, hotvet2, uotvet, uotvet2;

        float Z = 0.25f, c, random_number, text_zadachi, k; // переменные для расчета

        private void materialRaisedButton4_Click(object sender, EventArgs e)
        {
            exportWord.Application wordapp = new exportWord.Application();
            wordapp.Visible = true;
            exportWord.Document worddoc;
            object wordobj = System.Reflection.Missing.Value;
            worddoc = wordapp.Documents.Add(ref wordobj);
            wordapp.Selection.TypeText(textBox5.Text);
            wordapp.Selection.TypeText(textBox6.Text);
            wordapp = null;
        }

        private void materialRaisedButton3_Click(object sender, EventArgs e)
        {
            exportWord.Application wordapp = new exportWord.Application();
            wordapp.Visible = true;
            exportWord.Document worddoc;
            object wordobj = System.Reflection.Missing.Value;
            worddoc = wordapp.Documents.Add(ref wordobj);
            wordapp.Selection.TypeText(textBox1.Text);
            wordapp.Selection.TypeText(textBox2.Text);
            wordapp = null;
        }

        int M = 16;
        string empty = "\r\n\r\n"; //пробел вниз на 2 строки
        string numberz = "Задача №"; //номер задачи
        Random random = new Random();//случайное число, по котором генерируется задача
        Random rd = new Random();//случайный вывод задачи

        // формула, по которой будет решаться задача
        String formula2 = "2Fe(ТВ) + 3/2О₂(Г) → Fe₂О₃(ТВ) + 817 кДж";
        String formula3 = "2Mg + O₂ → 2MgO + 1204 кДж";
        String formula5 = "3Fe₃0₄(ТВ) + 8Al(ТВ) → 9Fe(ТВ) + 4Al₂O(ТВ) + 3330 кДж";
        String formula9 = "2Аl(ТВ) + 3/2О₂(Г) → Аl₂О₃(ТВ) + 1676 кДж";
        String formula15 = "4FeS₂(ТВ) + 11O₂(Г) → 8SO₂(Г) + 2Fe₂O₃ + 3310 кДж";
        String formula16 = "N₂(Г) + O₂(Г) → 2NO(Г) + 180 кДж";
        String formula1 = "CH₄ + 2O₂ → CO₂ + 2H₂O + 900 кДж";
        String formula4 = "2C₂H₂ + 5O₂ → 4CO₂ + 2H₂O + 2610 кДж";
        String formula7 = "C + O₂ → CO₂ + 402,24 кДж";
        String formula8 = "2С₂Н₂(Г) + 5О₂(Г) → 4СО₂(Г)+ 2Н₂О(Г) + 2602,4 кДж";
        String formula10 = "2С₃Н₈(Г) + 10О₂(Г) → 6СО₂(Г) + 8Н₂О(Ж) + 4440 кДж";
        String formula11 = "2KClO₃ → 2KCl + 3O₂ + 91 кДж";
        String formula12 = "H₂ + Cl₂ → 2HCl + 184,36 кДж";
        String formula13 = "C + O₂ → CO₂ + 402,24 кДж";

        String[] name; //массив на вывод задачи с ответом
        int zd;
        int i, j, d;


        private void materialRaisedButton2_Click(object sender, EventArgs e) //неорганика
        {
            string s = textBox4.Text;
            int count = Convert.ToInt32(s);

            if (materialRadioButton4.Checked) //случайно
            {
                textBox5.Clear();
                textBox6.Clear();

                for (i = 1; i < count + 1; i++)
                {
                    k = random.Next(10, 500);
                    random_number = k / 100; // эта и предыдущая строки генерируют случаное число от 0.1 до 5 с шагом 0.01, которое задействуется для генерации самой задачи
                    c = random_number * Z;
                    text_zadachi = c * M;

                    zotvet = new Decimal((text_zadachi * 817) / (2 * 55.85));
                    zotvet2 = Math.Round(zotvet, 2, MidpointRounding.AwayFromZero);
                    String zotvet3 = zotvet2 + "";

                    mgotvet = new Decimal(text_zadachi * 15.025);
                    mgotvet2 = Math.Round(mgotvet, 2, MidpointRounding.AwayFromZero);
                    String mgotvet3 = mgotvet2 + "";

                    alotvet = new Decimal(text_zadachi * 8.16176471);
                    alotvet2 = Math.Round(alotvet, 2, MidpointRounding.AwayFromZero);
                    String alotvet3 = alotvet2 + "";

                    al_2_otvet = new Decimal(text_zadachi * 16.4313725);
                    al_2_otvet2 = Math.Round(al_2_otvet, 2, MidpointRounding.AwayFromZero);
                    String al_2_otvet3 = al_2_otvet2 + "";

                    fesotvet = new Decimal(text_zadachi * 8 / 3310 * 22.4);
                    fesotvet2 = Math.Round(fesotvet, 2, MidpointRounding.AwayFromZero);
                    String fesotvet3 = fesotvet2 + "";

                    notvet = new Decimal(text_zadachi * 1 / 180 * 22.4);
                    notvet2 = Math.Round(fesotvet, 2, MidpointRounding.AwayFromZero);
                    String notvet3 = notvet2 + "";
                    //текст задачи
                    name = new String[6];
                    name[0] = (formula2) + "\r\n\r\nОпределите тепловой эффект реакции, в которой из " + text_zadachi + "г железа и необходимого количества кислорода образовался оксид железа(III).";
                    name[1] = (formula3) + "\r\n\r\nОпределите количество теплоты, которое выделится при образовании " + text_zadachi + " г MgO  в результате реакции горения магния, с помощью термохимического уравнения.";
                    name[2] = (formula5) + "\r\n\r\nОпределите тепловой эффект реакции, в которой из " + text_zadachi + "г алюминия и необходимого количества кислорода образовался оксид алюминия.";
                    name[3] = (formula9) + "\r\n\r\nОпределите тепловой эффект реакции, в которой из " + text_zadachi + "г алюминия и необходимого количества кислорода образовался оксид алюминия.";
                    name[4] = (formula15) + "\r\n\r\nВ результате реакции выделилось " + text_zadachi + "кДж теплоты. Определите объем(л) выделившегося диоксида серы(н.у.)";
                    name[5] = (formula16) + "\r\n\r\nКакой объем азота нужно сжечь, чтобы поглотилось " + text_zadachi + " тепла?";

                    zd = rd.Next(6);
                    textBox5.Text += numberz + i + "\r\n\r\n" + name[zd] + empty;
                    if (zd == 0)
                    {
                        textBox6.Text += i + ". Ответ: " + (zotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 1)
                    {
                        textBox6.Text += i + ". Ответ: " + (mgotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 2)
                    {
                        textBox6.Text += i + ". Ответ: " + (alotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 3)
                    {
                        textBox6.Text += i + ". Ответ: " + (al_2_otvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 4)
                    {
                        textBox6.Text += i + ". Ответ: " + (fesotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 5)
                    {
                        textBox6.Text += i + ". Ответ: " + (notvet3) + " кДж." + "\r\n";
                    }
                }
            }

            else if (materialRadioButton5.Checked) //кол-во теплоты. Блок закончен
            {
                textBox5.Clear();
                textBox6.Clear();

                for (j = 1; j < count + 1; j++)
                {
                    k = random.Next(10, 500);
                    random_number = k / 100; // эта и предыдущая строки генерируют случаное число от 0.1 до 5 с шагом 0.01, которое задействуется для генерации самой задачи
                    c = random_number * Z;
                    text_zadachi = c * M;

                    zotvet = new Decimal((text_zadachi * 817) / (2 * 55.85));
                    zotvet2 = Math.Round(zotvet, 2, MidpointRounding.AwayFromZero);
                    String zotvet3 = zotvet2 + "";

                    mgotvet = new Decimal(text_zadachi * 15.025);
                    mgotvet2 = Math.Round(mgotvet, 2, MidpointRounding.AwayFromZero);
                    String mgotvet3 = mgotvet2 + "";

                    alotvet = new Decimal(text_zadachi * 8.16176471);
                    alotvet2 = Math.Round(alotvet, 2, MidpointRounding.AwayFromZero);
                    String alotvet3 = alotvet2 + "";

                    al_2_otvet = new Decimal(text_zadachi * 16.4313725);
                    al_2_otvet2 = Math.Round(al_2_otvet, 2, MidpointRounding.AwayFromZero);
                    String al_2_otvet3 = al_2_otvet2 + "";
                    //текст задачи с ответом
                    name = new String[4];
                    name[0] = (formula2) + "\r\n\r\nОпределите тепловой эффект реакции, в которой из " + text_zadachi + "г железа и необходимого количества кислорода образовался оксид железа(III).";
                    name[1] = (formula3) + "\r\n\r\nОпределите количество теплоты, которое выделится при образовании " + text_zadachi + " г MgO  в результате реакции горения магния, с помощью термохимического уравнения.";
                    name[2] = (formula5) + "\r\n\r\nОпределите тепловой эффект реакции, в которой из " + text_zadachi + "г алюминия и необходимого количества кислорода образовался оксид алюминия.";
                    name[3] = (formula9) + "\r\n\r\nОпределите тепловой эффект реакции, в которой из " + text_zadachi + "г алюминия и необходимого количества кислорода образовался оксид алюминия.";

                    zd = rd.Next(4);
                    textBox5.Text += numberz + j + "\r\n\r\n" + name[zd] + empty;
                    if (zd == 0)
                    {
                        textBox6.Text += j + ". Ответ: " + (zotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 1)
                    {
                        textBox6.Text += j + ". Ответ: " + (mgotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 2)
                    {
                        textBox6.Text += j + ". Ответ: " + (alotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 3)
                    {
                        textBox6.Text += j + ". Ответ: " + (al_2_otvet3) + " кДж." + "\r\n";
                    }
                }
            }
            else if (materialRadioButton6.Checked) //масса вещества
            {
                textBox5.Clear();
                textBox6.Clear();

                for (d = 1; d < count + 1; d++)
                {
                    k = random.Next(10, 500);
                    random_number = k / 100; // эта и предыдущая строки генерируют случаное число от 0.1 до 5 с шагом 0.01, которое задействуется для генерации самой задачи
                    c = random_number * Z;
                    text_zadachi = c * M;

                    fesotvet = new Decimal(text_zadachi * 8 / 3310 * 22.4);
                    fesotvet2 = Math.Round(fesotvet, 2, MidpointRounding.AwayFromZero);
                    String fesotvet3 = fesotvet2 + "";

                    notvet = new Decimal(text_zadachi * 1 / 180 * 22.4);
                    notvet2 = Math.Round(fesotvet, 2, MidpointRounding.AwayFromZero);
                    String notvet3 = notvet2 + "";

                    //текст задачи
                    name = new String[2];
                    name[0] = (formula15) + "\r\n\r\nВ результате реакции выделилось " + text_zadachi + "кДж теплоты. Определите объем(л) выделившегося диоксида серы(н.у.)";
                    name[1] = (formula16) + "\r\n\r\nКакой объем азота нужно сжечь, чтобы поглотилось " + text_zadachi + " тепла?";

                    zd = rd.Next(2);
                    textBox5.Text += numberz + d + "\r\n\r\n" + name[zd] + empty;
                    if (zd == 0)
                    {
                        textBox6.Text += d + ". Ответ: " + (fesotvet3) + " л" + "\r\n";
                    }
                    if (zd == 1)
                    {
                        textBox6.Text += d + ". Ответ: " + (notvet3) + " л" + "\r\n";
                    }
                }
            }
        }                 

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void materialTabSelector1_Click(object sender, EventArgs e)
        {
            
        }

        private void materialRaisedButton1_Click(object sender, EventArgs e) //органика
        {
            string s = textBox3.Text;
            int count = Convert.ToInt32(s);

            textBox1.Clear();
            textBox2.Clear();

            if (materialRadioButton1.Checked)
            {
                for (i = 1; i < count + 1; i++)
                {
                    k = random.Next(10, 500);
                    random_number = k / 100; // эта и предыдущая строки генерируют случаное число от 0.1 до 5 с шагом 0.01, которое задействуется для генерации самой задачи
                    c = random_number * Z;
                    text_zadachi = c * M;

                    //здесь происходит расчет задач
                    otvet = new Decimal(c * 900 / 1);
                    otvet2 = Math.Round(otvet, 2, MidpointRounding.AwayFromZero);
                    String otvet3 = otvet2 + "";

                    acotvet = new Decimal(text_zadachi * 50.1923077);
                    acotvet2 = Math.Round(acotvet, 2, MidpointRounding.AwayFromZero);
                    String acotvet3 = acotvet2 + "";

                    cotvet = new Decimal(text_zadachi * 33.52);
                    cotvet2 = Math.Round(cotvet, 2, MidpointRounding.AwayFromZero);
                    String cotvet3 = cotvet2 + "";

                    ac_2_otvet = new Decimal(text_zadachi * 0.0580892857);
                    ac_2_otvet2 = Math.Round(ac_2_otvet, 2, MidpointRounding.AwayFromZero);
                    String ac_2_otvet3 = ac_2_otvet2 + "";

                    potvet = new Decimal(text_zadachi * 99.1071429);
                    potvet2 = Math.Round(potvet, 2, MidpointRounding.AwayFromZero);
                    String potvet3 = potvet2 + "";

                    ootvet = new Decimal(text_zadachi * 3 / 91 * 22.4);
                    ootvet2 = Math.Round(ootvet, 2, MidpointRounding.AwayFromZero);
                    String ootvet3 = ootvet2 + "";

                    hotvet = new Decimal(text_zadachi * 2 / 184.36 * 22.4);
                    hotvet2 = Math.Round(hotvet, 2, MidpointRounding.AwayFromZero);
                    String hotvet3 = hotvet2 + "";

                    uotvet = new Decimal(text_zadachi * 12 / 402.24);
                    uotvet2 = Math.Round(uotvet, 2, MidpointRounding.AwayFromZero);
                    String uotvet3 = uotvet2 + "";

                    //текст задачи с ответом                   
                    name = new String[8];
                    name[0] = (formula1) + "\r\n\r\nПо термохимическому уравнению горения метана определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "г метана.";
                    name[1] = (formula4) + "\r\n\r\nПо термохимическому уравнению горения ацетилена определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "г ацетилена.";
                    name[2] = (formula7) + "\r\n\r\nПо термохимическому уравнению горения метана определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "г метана.";
                    name[3] = (formula8) + "\r\n\r\nПо термохимическому уравнению горения ацетилена определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "мл ацетилена.";
                    name[4] = (formula10) + "\r\n\r\nПо термохимическому уравнению горения пропана определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "л пропана.";
                    name[5] = (formula11) + "\r\n\r\nКакой объем кислорода (при н.у.) выделится в результате реакции, если на разложение бертолетовой  соли было  затрачено.";
                    name[6] = (formula12) + "\r\n\r\nПо термохимическому уравнению рассчитайте, какой объем затрачен на образование хлороводорода (при н.у.), если при этом выделилось.";
                    name[7] = (formula13) + "\r\n\r\nОпределить кол-во сгоревшего угля, если в ходе горения было выделено  " + text_zadachi + "  кДж энергии.";

                    zd = rd.Next(8);
                    textBox1.Text += numberz + i + "\r\n\r\n" + name[zd] + empty;
                    if (zd == 0)
                    {
                        textBox2.Text += i + ". Ответ: " + (otvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 1)
                    {
                        textBox2.Text += i + ". Ответ: " + (acotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 2)
                    {
                        textBox2.Text += i + ". Ответ: " + (cotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 3)
                    {
                        textBox2.Text += i + ". Ответ: " + (ac_2_otvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 4)
                    {
                        textBox2.Text += i + ". Ответ: " + (potvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 5)
                    {
                        textBox2.Text += i + ". Ответ: " + (ootvet3) + " л." + "\r\n";
                    }
                    if (zd == 6)
                    {
                        textBox2.Text += i + ". Ответ: " + (hotvet3) + " л." + "\r\n";
                    }
                    if (zd == 7)
                    {
                        textBox2.Text += i + ". Ответ: " + (uotvet3) + " г." + "\r\n";
                    }
                }
            }
            else if (materialRadioButton2.Checked) //кол-во теплоты
            {
                textBox1.Clear();
                textBox2.Clear();

                for (j = 1; j < count + 1; j++)
                {
                    k = random.Next(10, 500);
                    random_number = k / 100; // эта и предыдущая строки генерируют случаное число от 0.1 до 5 с шагом 0.01, которое задействуется для генерации самой задачи
                    c = random_number * Z;
                    text_zadachi = c * M;

                    //здесь происходит расчет задач
                    otvet = new Decimal(c * 900 / 1);
                    otvet2 = Math.Round(otvet, 2, MidpointRounding.AwayFromZero);
                    String otvet3 = otvet2 + "";

                    acotvet = new Decimal(text_zadachi * 50.1923077);
                    acotvet2 = Math.Round(acotvet, 2, MidpointRounding.AwayFromZero);
                    String acotvet3 = acotvet2 + "";

                    cotvet = new Decimal(text_zadachi * 33.52);
                    cotvet2 = Math.Round(cotvet, 2, MidpointRounding.AwayFromZero);
                    String cotvet3 = cotvet2 + "";

                    ac_2_otvet = new Decimal(text_zadachi * 0.0580892857);
                    ac_2_otvet2 = Math.Round(ac_2_otvet, 2, MidpointRounding.AwayFromZero);
                    String ac_2_otvet3 = ac_2_otvet2 + "";

                    potvet = new Decimal(text_zadachi * 99.1071429);
                    potvet2 = Math.Round(potvet, 2, MidpointRounding.AwayFromZero);
                    String potvet3 = potvet2 + "";

                    //текст задачи с ответом
                    name = new String[5];
                    name[0] = (formula1) + "\r\n\r\nПо термохимическому уравнению горения метана определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "г метана.";
                    name[1] = (formula4) + "\r\n\r\nПо термохимическому уравнению горения ацетилена определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "г ацетилена.";
                    name[2] = (formula7) + "\r\n\r\nПо термохимическому уравнению горения метана определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "г метана.";
                    name[3] = (formula8) + "\r\n\r\nПо термохимическому уравнению горения ацетилена определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "мл ацетилена.";
                    name[4] = (formula10) + "\r\n\r\nПо термохимическому уравнению горения пропана определите, сколько выделиться теплоты, если сгорело " + text_zadachi + "л пропана.";

                    zd = rd.Next(5);
                    textBox1.Text += numberz + j + "\r\n\r\n" + name[zd] + empty;
                    if (zd == 0)
                    {
                        textBox2.Text += j + ". Ответ: " + (otvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 1)
                    {
                        textBox2.Text += j + ". Ответ: " + (acotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 2)
                    {
                        textBox2.Text += j + ". Ответ: " + (cotvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 3)
                    {
                        textBox2.Text += j + ". Ответ: " + (ac_2_otvet3) + " кДж." + "\r\n";
                    }
                    if (zd == 4)
                    {
                        textBox2.Text += j + ". Ответ: " + (potvet3) + " кДж." + "\r\n";
                    }
                }
            }
            else if (materialRadioButton3.Checked) //масса вещества
            {
                textBox1.Clear();
                textBox2.Clear();

                for (d = 1; d < count + 1; d++)
                {
                    k = random.Next(10, 500);
                    random_number = k / 100; // эта и предыдущая строки генерируют случаное число от 0.1 до 5 с шагом 0.01, которое задействуется для генерации самой задачи
                    c = random_number * Z;
                    text_zadachi = c * M;

                    ootvet = new Decimal(text_zadachi * 3 / 91 * 22.4);
                    ootvet2 = Math.Round(ootvet, 2, MidpointRounding.AwayFromZero);
                    String ootvet3 = ootvet2 + "";

                    hotvet = new Decimal(text_zadachi * 2 / 184.36 * 22.4);
                    hotvet2 = Math.Round(hotvet, 2, MidpointRounding.AwayFromZero);
                    String hotvet3 = hotvet2 + "";

                    uotvet = new Decimal(text_zadachi * 12 / 402.24);
                    uotvet2 = Math.Round(uotvet, 2, MidpointRounding.AwayFromZero);
                    String uotvet3 = uotvet2 + "";

                    //текст задачи с ответом
                    name = new String[3];
                    name[0] = (formula11) + "\r\n\r\nКакой объем кислорода (при н.у.) выделится в результате реакции, если на разложение бертолетовой  соли было  затрачено " + text_zadachi + " кДж теплоты.";
                    name[1] = (formula12) + "\r\n\r\nПо термохимическому уравнению рассчитайте, какой объем затрачен на образование хлороводорода (при н.у.), если при этом выделилось " + text_zadachi + "кДж теплоты.";
                    name[2] = (formula13) + "\r\n\r\nОпределить кол-во сгоревшего угля, если в ходе горения было выделено  " + text_zadachi + "  кДж энергии.";

                    zd = rd.Next(3);
                    textBox1.Text += numberz + d + "\r\n\r\n" + name[zd] + empty;
                    if (zd == 0)
                    {
                        textBox2.Text += d + ". Ответ: " + (ootvet3) + " л" + "\r\n";
                    }
                    if (zd == 1)
                    {
                        textBox2.Text += d + ". Ответ: " + (hotvet3) + " л" + "\r\n";
                    }
                    if (zd == 2)
                    {
                        textBox2.Text += d + ". Ответ: " + (uotvet3) + " г" + "\r\n";
                    }
                }
            }
        }
    }
}               