using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Ports;
using System.Runtime.InteropServices;
using OMICRON.CMEngAL;
using System.Globalization;
using System.ComponentModel;
using System.Threading;

namespace OmicronController
{
    public class OmicronControllerClass
    {
        static int selectedDeviceId = -1;
        static ICMEngine engine;


        static double Voltage1 = 230;
        static double Voltage2 = 230;
        static double Voltage3 = 230;


        static double angle1 = 0;
        static double angle2 = 240;
        static double angle3 = 120;


        static double angle1_v_original = 0;
        static double angle2_v_original = 240;
        static double angle3_v_original = 120;

        static double current1 = 5;
        static double current2 = 5;
        static double current3 = 5;

        static double anglec1 = 0;
        static double anglec2 = 240;
        static double anglec3 = 120;

        static double anglec1_original = 0;
        static double anglec2_original = 240;
        static double anglec3_original = 120;


        static bool current1_on = false;
        static bool current2_on = false;
        static bool current3_on = false;

        static double frequency = 50.0;

        static bool dosage_on = false;
        static bool dosage_running = false;

        static int energia_dosage = 0;
        static int energia_ja_injectada = 0;

        static DateTime inicio = DateTime.Now;
        static DateTime fim = DateTime.Now;
        static DateTime atual = DateTime.Now;




        public static void OmicronController()
        {
            engine = new CMEngine();
            Initialize();
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;
        }



        public static void Initialize()
        {
            string result;
            engine.DevScanForNew(false);
            result = engine.DevGetList(ListSelectType.lsUnlockedAssociated);
            // select first device if one is available
            if (result == null)
            {
                Console.WriteLine("No device available!");

                return;
            }
            Console.WriteLine(result);
            if (!int.TryParse(result.Substring(0, result.IndexOf(',')), out selectedDeviceId))
            {
                Console.WriteLine("Can't select device.");
                return;
            }
            selectedDeviceId = 1;
            Console.WriteLine(string.Format("Device with ID {0} selected ({1}).", selectedDeviceId, engine.get_SerialNumber(selectedDeviceId)));

            engine.DevLock(selectedDeviceId);
            Console.WriteLine(engine.Exec(selectedDeviceId, "sys:reset"));

            result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);

            result = engine.Exec(selectedDeviceId, "out:nomfreq(50.000000, int)");
            Console.WriteLine("System status: " + result);

            result = engine.Exec(selectedDeviceId, "amp: def(6, int)");
            Console.WriteLine("System status: " + result);

            result = engine.Exec(selectedDeviceId, "amp: route(clrnooff)");
            Console.WriteLine("System status: " + result);

            result = engine.Exec(selectedDeviceId, "amp: def(clrnooff)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: route(v(1), 9)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: route(i(1), 17)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: ctrl(i(1), 15.000000)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: def(off)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: highpwr(int, i, off)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: highpwr(int, v, off)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "cb: free");
            Console.WriteLine("System status: " + result);

            result = engine.Exec(selectedDeviceId, "amp:range(v(1), 288.675)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: range(i(1), 7)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "amp: ovl(0.050000)");
            Console.WriteLine("System status: " + result);
            result = engine.Exec(selectedDeviceId, "out:ana: pmode(abs)");
            Console.WriteLine("System status: " + result);

        }


        public static void Stop()
        {
            engine.DevUnlock(selectedDeviceId);
        }


        public static void VoltageSupplyState(bool on)
        {
            if (on == false)
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1):off"));
                Console.WriteLine("Omicron off");
            }
            else
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1):on"));

                Console.WriteLine("Omicron on");
            }

            string result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
        }


        public static void CurrentSupplyState(bool on1, bool on2, bool on3)
        {
            if (on1 == false)
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:i(1:1):off"));
            }
            else
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:i(1:1):on"));
            }

            if (on2 == false)
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:i(1:2):off"));
            }
            else
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:i(1:2):on"));
            }

            if (on3 == false)
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:i(1:3):off"));
            }
            else
            {
                Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:i(1:3):on"));
            }

            string result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
        }


        public static void ChangeVoltage(double V1, double V2, double V3, double ANG1, double ANG2, double ANG3, double FREQ)
        {
            Voltage1 = V1;
            Voltage2 = V2;
            Voltage3 = V3;


            angle1 = ANG1;
            angle2 = ANG2;
            angle3 = ANG3;


            angle1_v_original = angle1;
            angle2_v_original = angle2;
            angle3_v_original = angle3;

            if (angle1 >= 360)
            {
                angle1 = angle1 - 360;
            }
            if (angle2 >= 360)
            {
                angle2 = angle2 - 360;
            }
            if (angle3 >= 360)
            {
                angle3 = angle3 - 360;
            }

            //angle1 = ANG1;
            //angle2 = ANG2;
            //angle3 = ANG3;





            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:1):a({0});f({1});p({2});wav(sin)", V1, FREQ, angle1)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:2):a({0});f({1});p({2});wav(sin)", V2, FREQ, angle2)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:3):a({0});f({1});p({2});wav(sin)", V3, FREQ, angle3)));


            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(1) ON: amplitude={0}; frequency={1}; anglec={2} ", V1, FREQ, angle1));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(2) ON: amplitude={0}; frequency={1}; anglec={2}", V2, FREQ, angle2));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(3) ON: amplitude={0}; frequency={1}; anglec={2}", V3, FREQ, angle3));


            string result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
        }

        public static void ChangeVoltage(double V1, double V2, double V3, double ANG1, double ANG2, double ANG3, double FREQ1, double FREQ2, double FREQ3)
        {
            Voltage1 = V1;
            Voltage2 = V2;
            Voltage3 = V3;


            angle1 = ANG1;
            angle2 = ANG2;
            angle3 = ANG3;


            angle1_v_original = angle1;
            angle2_v_original = angle2;
            angle3_v_original = angle3;

            if (angle1 >= 360)
            {
                angle1 = angle1 - 360;
            }
            if (angle2 >= 360)
            {
                angle2 = angle2 - 360;
            }
            if (angle3 >= 360)
            {
                angle3 = angle3 - 360;
            }

            //angle1 = ANG1;
            //angle2 = ANG2;
            //angle3 = ANG3;





            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:1):a({0});f({1});p({2});wav(sin)", V1, FREQ1, angle1)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:2):a({0});f({1});p({2});wav(sin)", V2, FREQ2, angle2)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:3):a({0});f({1});p({2});wav(sin)", V3, FREQ3, angle3)));


            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(1) ON: amplitude={0}; frequency={1}; anglec={2} ", V1, FREQ1, angle1));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(2) ON: amplitude={0}; frequency={1}; anglec={2}", V2, FREQ1, angle2));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(3) ON: amplitude={0}; frequency={1}; anglec={2}", V3, FREQ1, angle3));


            string result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
        }

        public static void ChangeVoltage(double V1, double V2, double V3, double ANG1, double ANG2, double ANG3, double FREQ1, double FREQ2, double FREQ3, List<double> harmonicNumber, List<double> harmonicPercentage, List<double> harmonicPhase)
        {
            Voltage1 = V1;
            Voltage2 = V2;
            Voltage3 = V3;


            angle1 = ANG1;
            angle2 = ANG2;
            angle3 = ANG3;


            angle1_v_original = angle1;
            angle2_v_original = angle2;
            angle3_v_original = angle3;

            if (angle1 >= 360)
            {
                angle1 = angle1 - 360;
            }
            if (angle2 >= 360)
            {
                angle2 = angle2 - 360;
            }
            if (angle3 >= 360)
            {
                angle3 = angle3 - 360;
            }

            //angle1 = ANG1;
            //angle2 = ANG2;
            //angle3 = ANG3;
            string harmonics = "";
            for (int i = 0; i < harmonicNumber.Count; i++)
            {
                harmonics += $"{harmonicNumber[i].ToString("0")},{harmonicPercentage[i].ToString("0.000000")},{harmonicPhase[i].ToString("0.000000")}";
                if (i < (harmonicNumber.Count - 1))
                {
                    harmonics += ",";
                }
            }

            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:1):a({0});f({1});p({2});wav(sumcorr,{1},0.0,{3})", V1, FREQ1, angle1, harmonics)));
            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:2):a({0});f({1});p({2});wav(sumcorr,{1},0.0,{3})", V2, FREQ2, angle2, harmonics)));
            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:3):a({0});f({1});p({2});wav(sumcorr,{1},0.0,{3})", V3, FREQ3, angle3, harmonics)));
            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:1):sig(1):wav(sin)"));

            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:1):sig(1):a({Voltage1.ToString("0.000000")});p({angle1.ToString("0.000000")})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:1):sig(1):f({FREQ1.ToString("0.000000")})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:1):sig(2):wav(sumcorr,{FREQ1.ToString("0.000000")},0.0,{harmonics})"));

            var teste = $"out:ana:v(1:1):sig(2):a({Voltage1.ToString("0.000000")});p({angle1.ToString("0.000000")})";
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:1):sig(2):a({Voltage1.ToString("0.000000")});p({angle1.ToString("0.000000")})"));


            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:2):sig(1):wav(sin)"));

            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:2):sig(1):a({Voltage2.ToString("0.000000")});p({angle2.ToString("0.000000")})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:2):sig(1):f({FREQ2.ToString("0.000000")})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:2):sig(2):wav(sumcorr,{FREQ2.ToString("0.000000")},0.0,{harmonics})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:2):sig(2):a({Voltage2.ToString("0.000000")});p({angle2.ToString("0.000000")})"));

            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:3):sig(1):wav(sin)"));

            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:3):sig(1):a({Voltage3.ToString("0.000000")});p({angle3.ToString("0.000000")})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:3):sig(1):f({FREQ3.ToString("0.000000")})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:3):sig(2):wav(sumcorr,{FREQ3.ToString("0.000000")},0.0,{harmonics})"));
            Console.WriteLine(engine.Exec(selectedDeviceId, $"out:ana:v(1:3):sig(2):a({Voltage3.ToString("0.000000")});p({angle3.ToString("0.000000")})"));

            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:1):mix(sig(1))"));
            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:2):mix(sig(1))"));
            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:3):mix(sig(1))"));


            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:1):mix(add)"));
            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:2):mix(add)"));
            Console.WriteLine(engine.Exec(selectedDeviceId, "out:ana:v(1:3):mix(add)"));

            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:1):a({0});f({1});p({2}):sig(2):wav(sumcorr,50.000000,0.0,3,0.100000,0.000000,5,0.000000,0.000000)", V1, FREQ1, angle1, harmonics)));

            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(1) ON: amplitude={0}; frequency={1}; anglec={2} ", V1, FREQ1, angle1));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(2) ON: amplitude={0}; frequency={1}; anglec={2}", V2, FREQ1, angle2));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(3) ON: amplitude={0}; frequency={1}; anglec={2}", V3, FREQ1, angle3));


            string result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
        }




        public static void ChangeCurrent(double I1, double I2, double I3, double ANG1, double ANG2, double ANG3, double FREQ)
        {
            current1 = I1;
            current2 = I2;
            current3 = I3;

            anglec1 = ANG1;
            anglec2 = ANG2;
            anglec3 = ANG3;

            anglec1_original = ANG1;
            anglec2_original = ANG2;
            anglec3_original = ANG3;


            if (anglec1 >= 360)
            {
                anglec1 = anglec1 - 360;
            }
            if (anglec2 >= 360)
            {
                anglec2 = anglec2 - 360;
            }
            if (anglec3 >= 360)
            {
                anglec3 = anglec3 - 360;
            }



            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:1):a({0});f({1});p({2});wav(sin)", I1, FREQ, anglec1)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:2):a({0});f({1});p({2});wav(sin)", I2, FREQ, anglec2)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:3):a({0});f({1});p({2});wav(sin)", I3, FREQ, anglec3)));


            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(1) ON: current={0}; frequency={1}; angle={2}", I1, FREQ, anglec1));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(2) ON: current={0}; frequency={1}; angle={2}", I2, FREQ, anglec2));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(3) ON: current={0}; frequency={1}; angle={2}", I3, FREQ, anglec3));

            string result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);


        }


        public static void Dosage(double Energy, double I1, double I2, double I3, double ANG1, double ANG2, double ANG3, double FREQ)
        {
            string result = "";

            current1 = I1;
            current2 = I2;
            current3 = I3;

            //anglec1 = 0;
            //anglec2 = 240.0;
            //anglec3 = 120.0;

            //anglec1_original = 0;
            //anglec2_original = 240.0;
            //anglec3_original = 120.0;

            //angle1 = ANG1;
            //angle2 = ANG2;
            //angle3 = ANG3;

            //angle1_v_original = angle1;
            //angle2_v_original = angle2;
            //angle3_v_original = angle3;

            //if (angle1 >= 360)
            //{
            //    angle1 = angle1 - 360;
            //}
            //if (angle2 >= 360)
            //{
            //    angle2 = angle2 - 360;
            //}
            //if (angle3 >= 360)
            //{
            //    angle3 = angle3 - 360;
            //}

            anglec1 = ANG1;
            anglec2 = ANG2;
            anglec3 = ANG3;

            anglec1_original = anglec1;
            anglec2_original = anglec2;
            anglec3_original = anglec3;

            if (anglec1 >= 360)
            {
                anglec1 = anglec1 - 360;
            }
            if (anglec2 >= 360)
            {
                anglec2 = anglec2 - 360;
            }
            if (anglec3 >= 360)
            {
                anglec3 = anglec3 - 360;
            }



            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:1):a({0});f({1});p({2});wav(sin)", I1, FREQ, anglec1)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:2):a({0});f({1});p({2});wav(sin)", I2, FREQ, anglec2)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:3):a({0});f({1});p({2});wav(sin)", I3, FREQ, anglec3)));


            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(1) ON: current={0}; frequency={1}; angle={2}", I1, FREQ, anglec1));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(2) ON: current={0}; frequency={1}; angle={2}", I2, FREQ, anglec2));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(3) ON: current={0}; frequency={1}; angle={2}", I3, FREQ, anglec3));


            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:1):a({0});f({1});p({2});wav(sin)", Voltage1, FREQ, angle1)));
            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:2):a({0});f({1});p({2});wav(sin)", Voltage2, FREQ, angle2)));
            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:3):a({0});f({1});p({2});wav(sin)", Voltage3, FREQ, angle3)));


            //Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(1) ON: amplitude={0}; frequency={1}; anglec={2} ", Voltage1, FREQ, angle1));
            //Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(2) ON: amplitude={0}; frequency={1}; anglec={2}", Voltage2, FREQ, angle2));
            //Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(3) ON: amplitude={0}; frequency={1}; anglec={2}", Voltage3, FREQ, angle3));

            inicio = DateTime.Now;

            result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
            CurrentSupplyState(true, true, true);

            while (true)
            {

                TimeSpan tempo_decorrido = DateTime.Now - inicio;

                energia_ja_injectada = (int)Math.Abs((tempo_decorrido.TotalSeconds / 3600.0 * (Math.Cos(Math.Abs(angle1_v_original - anglec1_original) * 3.1415 / 180.0) * Voltage1 * current1 + Math.Cos(Math.Abs(angle2_v_original - anglec2_original) * 3.1415 / 180.0) * Voltage2 * current2 + Math.Cos(Math.Abs(angle3_v_original - anglec3_original) * 3.1415 / 180.0) * Voltage3 * current3)));
                Console.WriteLine("Energy: " + energia_ja_injectada.ToString());


                if (energia_ja_injectada >= Energy)
                {
                    dosage_on = true;
                    dosage_running = false;

                    engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:1):a({0});f({1});p({2});wav(sin)", 0, frequency, anglec1));
                    engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:2):a({0});f({1});p({2});wav(sin)", 0, frequency, anglec2));
                    engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:3):a({0});f({1});p({2});wav(sin)", 0, frequency, anglec3));


                    //     engine.Exec(selectedDeviceId, "out:ana:i(1):off");
                    current1_on = false;
                    current2_on = false;
                    current3_on = false;

                    CurrentSupplyState(false, false, false);
                    break;
                }

                result = engine.Exec(selectedDeviceId, "sys:status?");
                Console.WriteLine("System status: " + result);

            }



            result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);




        }


        public static void DosageReactive(double Energy, double I1, double I2, double I3, double ANG1, double ANG2, double ANG3, double FREQ)
        {
            string result = "";

            current1 = I1;
            current2 = I2;
            current3 = I3;

            //anglec1 = 0;
            //anglec2 = 240.0;
            //anglec3 = 120.0;

            //anglec1_original = 0;
            //anglec2_original = 240.0;
            //anglec3_original = 120.0;

            //angle1 = ANG1;
            //angle2 = ANG2;
            //angle3 = ANG3;

            //angle1_v_original = angle1;
            //angle2_v_original = angle2;
            //angle3_v_original = angle3;

            //if (angle1 >= 360)
            //{
            //    angle1 = angle1 - 360;
            //}
            //if (angle2 >= 360)
            //{
            //    angle2 = angle2 - 360;
            //}
            //if (angle3 >= 360)
            //{
            //    angle3 = angle3 - 360;
            //}


            anglec1 = ANG1;
            anglec2 = ANG2;
            anglec3 = ANG3;

            anglec1_original = anglec1;
            anglec2_original = anglec2;
            anglec3_original = anglec3;

            if (anglec1 >= 360)
            {
                anglec1 = anglec1 - 360;
            }
            if (anglec2 >= 360)
            {
                anglec2 = anglec2 - 360;
            }
            if (anglec3 >= 360)
            {
                anglec3 = anglec3 - 360;
            }



            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:1):a({0});f({1});p({2});wav(sin)", I1, FREQ, anglec1)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:2):a({0});f({1});p({2});wav(sin)", I2, FREQ, anglec2)));
            Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:3):a({0});f({1});p({2});wav(sin)", I3, FREQ, anglec3)));


            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(1) ON: current={0}; frequency={1}; angle={2}", I1, FREQ, anglec1));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(2) ON: current={0}; frequency={1}; angle={2}", I2, FREQ, anglec2));
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "I(3) ON: current={0}; frequency={1}; angle={2}", I3, FREQ, anglec3));

            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:1):a({0});f({1});p({2});wav(sin)", Voltage1, FREQ, angle1)));
            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:2):a({0});f({1});p({2});wav(sin)", Voltage2, FREQ, angle2)));
            //Console.WriteLine(engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:v(1:3):a({0});f({1});p({2});wav(sin)", Voltage3, FREQ, angle3)));


            //Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(1) ON: amplitude={0}; frequency={1}; anglec={2} ", Voltage1, FREQ, angle1));
            //Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(2) ON: amplitude={0}; frequency={1}; anglec={2}", Voltage2, FREQ, angle2));
            //Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "V(3) ON: amplitude={0}; frequency={1}; anglec={2}", Voltage3, FREQ, angle3));

            inicio = DateTime.Now;

            result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
            CurrentSupplyState(true, true, true);

            while (true)
            {

                TimeSpan tempo_decorrido = DateTime.Now - inicio;

                energia_ja_injectada = (int)Math.Abs((tempo_decorrido.TotalSeconds / 3600.0 * (Math.Sin(Math.Abs(angle1_v_original - anglec1_original) * 3.1415 / 180.0) * Voltage1 * current1 + Math.Sin(Math.Abs(angle2_v_original - anglec2_original) * 3.1415 / 180.0) * Voltage2 * current2 + Math.Sin(Math.Abs(angle3_v_original - anglec3_original) * 3.1415 / 180.0) * Voltage3 * current3)));
                Console.WriteLine("Energy: " + energia_ja_injectada.ToString());


                if (energia_ja_injectada >= Energy)
                {
                    dosage_on = true;
                    dosage_running = false;

                    engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:1):a({0});f({1});p({2});wav(sin)", 0, frequency, anglec1));
                    engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:2):a({0});f({1});p({2});wav(sin)", 0, frequency, anglec2));
                    engine.Exec(selectedDeviceId, string.Format(CultureInfo.InvariantCulture, "out:ana:i(1:3):a({0});f({1});p({2});wav(sin)", 0, frequency, anglec3));


                    //     engine.Exec(selectedDeviceId, "out:ana:i(1):off");
                    current1_on = false;
                    current2_on = false;
                    current3_on = false;

                    CurrentSupplyState(false, false, false);
                    break;
                }

                result = engine.Exec(selectedDeviceId, "sys:status?");
                Console.WriteLine("System status: " + result);

            }
            result = engine.Exec(selectedDeviceId, "sys:status?");
            Console.WriteLine("System status: " + result);
        }
    }
}
