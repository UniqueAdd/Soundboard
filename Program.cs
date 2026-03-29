using System;
using System.Collections.Generic;
using System.Linq;
using NAudio.CoreAudioApi;

class Program
{
    private static readonly String DEFAULT_AUDIO_CABLE = "VB-Audio Virtual Cable";
    private static readonly String DEFAULT_APPLICATION = "firefox";


    static void Run()
    {
        //todo remove the DEFAULT_AUDIO_CABLE here and replace it with: "your virtual audio cable name here"
        String audioCable = DEFAULT_AUDIO_CABLE;
        //todo remove the DEFAULT_APPLICATION here and replace it with: "application you want here", "annother application if you want", "you can have as many as you want"
        List<String> applications = new List<String> {DEFAULT_APPLICATION};

        MMDeviceEnumerator enumerator = new MMDeviceEnumerator();
        //default speaker input as listed in system settings
        try
        {
            Setup.SetMicInput(enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred when trying to setup Microphon Input: " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
        //default microphone input as listed in system settings
        try 
        {
            Setup.SetHardOutput(enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia));
        }
        catch(Exception ex)
        {
            Console.WriteLine("An error occurred When trying to set Hardware Output: " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
        //todo you can change the Mic/Speaker/Headphone input/output here if you'd like. you just have to uncomment what's bellow this
        //Setup.SetMicInput(enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.All).FirstOrDefault(d => d.FriendlyName.Contains("Your Mic here")));
        //Setup.SetHardOutput(enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.All).FirstOrDefault(d => d.FriendlyName.Contains("Your Speaker/Headphone here")));

        //Sets any virtual audio cable you have here
        MMDevice? tempCable = null;
        foreach (MMDevice device in enumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active))
        {
            if (device.FriendlyName.Contains(audioCable))
            {
                tempCable = device;
                break;
            }
        }
        if (tempCable == null)
        {
            Console.WriteLine("Virtual Audio Cable does not exits. please change the device to one which exists.");
            Environment.Exit(1);
        }
        else
        {
            Setup.SetVirtualCable(tempCable);
        }
        //Setups all the applications you want to send to the virual audio cable here.
        try
        {
            Setup.SetListeningApplications(applications);
        }
        catch(Exception ex)
        {
            Console.WriteLine("An error occurred when trying to set applications to Listen to: " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
        //Setup the mixer and outputs to be ready to start the program.
        try 
        {
            Setup.Initialize();
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred when initializing : " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
        //litterally starts the program, starts the capturing and outputting of audio.
        try 
        {
            Setup.Start();
        }
        catch(Exception ex)
        {
            Console.WriteLine("An error occurred when starting Soundboard: " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }
    //Main Method.
    static void Main(String[] args)
    {
        //upon closing the program; run OnClose or OnCancel depending on what you do.
        AppDomain.CurrentDomain.ProcessExit += OnClose;
        Console.CancelKeyPress += OnCancel;
        //AAAAAHHHHHHHHHH IT CRASHED WHAT DO I DO
        AppDomain.CurrentDomain.UnhandledException += OnCrash;

        
        //App is running.
        Console.WriteLine("I AM ALIVE!!!\nFEAR ME MORTAL!\n");

        try
        {
            Run();
            Console.WriteLine("Audio routed. Press Ctrl+C or close to stop.\n");
            char key = Console.ReadKey().KeyChar;
            while (key.Equals(ConsoleSpecialKey.ControlC) == false)
            {
                key = Console.ReadKey().KeyChar;
                //runs program endlessly until user presses ctrl+c or closes the program.
            }
            Setup.Stop();
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }


    //on close stop audio capturing and reset outputs.
    static void OnClose(object sender, EventArgs e)
    {
        Console.WriteLine("Annihilation in progress...\n");
        Setup.Stop();
        Console.WriteLine("Annihilation Complete!");
        Environment.Exit(0);
    }
    //on cancel stop audio capturing and reset outputs. forces closing of program.
    static void OnCancel(object sender, ConsoleCancelEventArgs e)
    {
        e.Cancel = true;
        Console.WriteLine("Annihilation in progress...\n");
        Setup.Stop();
        Console.WriteLine("Annihilation Complete! Press any key to kill\n");
        Console.ReadKey();
        Environment.Exit(0);
    }
    //in case of crash
    public static void OnCrash(object sender, UnhandledExceptionEventArgs e)
    {
        Console.WriteLine("An unhandled exception occurred\n");
        Console.WriteLine("Annihilation in progress...\n");
        Setup.Stop();
        Console.WriteLine("Annihilation Complete! press any key to kill\n");
        Console.ReadKey();
        Environment.Exit(1);
    }
}