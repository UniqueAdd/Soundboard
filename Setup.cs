using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

public static class Setup
{
    //All variables used for the soundboard app
    private static WasapiCapture? micCapture = null;
    private static BufferedWaveProvider? micBuffer = null;
    private static WasapiCapture? cableCapure = null;
    private static BufferedWaveProvider? cableBuffer = null;
    private static MMDevice? virtualCableDevice = null;
    private static List<String>? applications = null;
    private static MMDevice? hardwareAudioDevice = null;
    private static WasapiCapture? hardCapture = null;
    private static BufferedWaveProvider? hardBuffer = null;
    private static MixingSampleProvider? virtualMixer = null;
    private static MixingSampleProvider? hardMixer = null;
    private static WasapiOut? virtualOutput = null;
    private static WasapiOut? hardOutput = null;
    
    // within .exe file, this function is called to set user mic input
    public static void SetMicInput(MMDevice mic)
    {
        try
        {
            micCapture = new WasapiCapture(mic);
            micBuffer = new BufferedWaveProvider(micCapture.WaveFormat);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error setting mic input: " + ex.Message);
        }
    }
    // within .exe file, this function is called to set user virtual audio cable input (ie. fireforx, youtube, google, etc.)
    public static void SetVirtualCable(MMDevice virtualCable)
    {
        try
        {
            virtualCableDevice = virtualCable;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error setting virtual cable: " + ex.Message);
        }
    }
    // within .exe file, this function is called to set what applications the audio cable send to both hardware output and virtual output.
    public static void SetListeningApplications(List<String> appList)
    {
        //moves all application given by User to outputs to audio cable.
        applications = appList;
        for (int i = 0; i < applications.Count; i++)
        {
            try
            {
                MoveToVirtualCable(applications[i]);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error moving application to virtual cable, application may not exist: " + ex.Message);
            }
        }
        try
        {
            cableCapure = new WasapiCapture(virtualCableDevice);
            cableBuffer = new BufferedWaveProvider(cableCapure.WaveFormat);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error setting up virtual cable capture: " + ex.Message);
        }
    }
    //function which moves the output of any application to the virtual audio cable. 
    private static void MoveToVirtualCable(String application)
    {
        var policyConfig = (IPolicyConfig)new PolicyConfigClient();
        try
        {
            //todo fix issue: Attempted to read or write protected memory. This is often an indication that other memory is corrupt. Repeat 2 times:
            var app = Process.GetProcessesByName(application)[0];
            policyConfig.SetPersistedDefaultAudioEndpoint(
            (uint)app.Id,
            DataFlow.Render,
            Role.Multimedia,
            virtualCableDevice.ID);
        }
        catch(Exception ex)
        {
            Console.WriteLine("An error occurred when trying to Find Application: " + ex.Message);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }
    // within .exe file, this function is called to set user hardware audio output (speakers, headphones, etc.)
    public static void SetHardOutput(MMDevice hard)
    {
        hardwareAudioDevice = hard;
        try 
        {
            hardCapture = new WasapiCapture(hard);
            hardBuffer = new BufferedWaveProvider(hardCapture.WaveFormat);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error setting hardware audio output: " + ex.Message);
        }
    }
    //setup for the mixer, which will combine the mic input, virtual audio cable input into one output to be sent
    //into other applications (discord, teams, etc.)
    //this works by shoving both the mic buffer and cable buffer into one mixer while also doing the same for 
    //the hard buffer, which is sent to the hardware output (speakers, headphones, etc.), and cable so both parties hear the sounds.
    private static void MixerSetup()
    {
        virtualMixer = new MixingSampleProvider(new []
        {
            micBuffer.ToSampleProvider(),
            cableBuffer.ToSampleProvider()
        });
        hardMixer = new MixingSampleProvider(new []
        {
            hardBuffer.ToSampleProvider(),
            cableBuffer.ToSampleProvider()
        });
    }
    //this function sets up the virtual and hardware outputs so that they are ready to use
    public static void Initialize()
    {
            if (micBuffer == null || cableBuffer == null || hardBuffer == null)
        {
            throw new InvalidOperationException("Oi Dummkopf. Must call SetMicInput and SetListeningApplications before Initialize.");
        }

        MixerSetup();
        //number at the end of constructor is latency. since this is virtual output, a lower latency is acceptable.
        //AudioClientShareMode is if other applications can also listen to the audio.
        //the false is if the audio is event driven (OS tells application when it needs more audio)
        //or timer driven (millisecond intervals).
        virtualOutput = new WasapiOut(virtualCableDevice, AudioClientShareMode.Shared, false, 50);
        virtualOutput.Init(virtualMixer);
        //number at the end of constructor is latency (Windows default is 10).
        hardOutput = new WasapiOut(hardwareAudioDevice, AudioClientShareMode.Shared, false, 7);
        hardOutput.Init(hardMixer);
    }
    //this runs the mic capture and everything-else capture to actually start the audio recording process.
    public static void Start()
    {
        micCapture.DataAvailable += (s, e) =>
            micBuffer.AddSamples(e.Buffer, 0, e.BytesRecorded);
        hardCapture.DataAvailable += (s, e) =>
            hardBuffer.AddSamples(e.Buffer, 0, e.BytesRecorded);
        cableCapure.DataAvailable += (s, e) =>
            cableBuffer.AddSamples(e.Buffer, 0, e.BytesRecorded);

        micCapture.StartRecording();
        hardCapture.StartRecording();
        cableCapure.StartRecording();
        virtualOutput.Play();
        hardOutput.Play();
    }
    //this function stops all recording and output, and moves all applications back to the default audio device.
    public static void Stop()
    {
        micCapture.StopRecording();
        hardCapture.StopRecording();
        cableCapure.StopRecording();
        virtualOutput.Stop();
        hardOutput.Stop();
        //disposal!!! YAY!!!!
        micCapture.Dispose();
        hardCapture.Dispose();
        cableCapure.Dispose();
        virtualOutput.Dispose();
        hardOutput.Dispose();
        virtualCableDevice.Dispose();

        for (int i = 0; i < applications.Count; i++)
        {
            try
            {
                MoveToDefaultAudioDevice(applications[i]);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error moving application back to default audio device, application may not exist: " + ex.Message);
            }
        }
    }
    //this function moves the output of any application back to the default audio device.
    private static void MoveToDefaultAudioDevice(String application)
    {
        var policyConfig = (IPolicyConfig)new PolicyConfigClient();
        var app = Process.GetProcessesByName(application)[0];
        policyConfig.SetPersistedDefaultAudioEndpoint(
            (uint)app.Id,
            DataFlow.Render,
            Role.Multimedia,
            hardwareAudioDevice.ID
        );
    }
}


//Policy Config
[Guid("f8679f50-850a-41cf-9c72-430f290290c8")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPolicyConfig
{
    [PreserveSig]
    int GetMixFormat(string pszDeviceName, IntPtr ppFormat);

    [PreserveSig]
    int GetDeviceFormat(string pszDeviceName, bool bDefault, IntPtr ppFormat);

    [PreserveSig]
    int ResetDeviceFormat(string pszDeviceName);

    [PreserveSig]
    int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);

    [PreserveSig]
    int GetProcessingPeriod(string pszDeviceName, bool bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);

    [PreserveSig]
    int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);

    [PreserveSig]
    int GetShareMode(string pszDeviceName, IntPtr pMode);

    [PreserveSig]
    int SetShareMode(string pszDeviceName, IntPtr mode);

    [PreserveSig]
    int GetPropertyValue(string pszDeviceName, bool bFxStore, IntPtr pKey, IntPtr pv);

    [PreserveSig]
    int SetPropertyValue(string pszDeviceName, bool bFxStore, IntPtr pKey, IntPtr pv);

    [PreserveSig]
    int SetDefaultEndpoint(string pszDeviceName, Role role);

    [PreserveSig]
    int SetPersistedDefaultAudioEndpoint(uint dwProcessId, DataFlow dataFlow, Role role, string pszDeviceName);

    [PreserveSig]
    int GetPersistedDefaultAudioEndpoint(uint dwProcessId, DataFlow dataFlow, Role role, out string pszDeviceName);

    [PreserveSig]
    int ClearAllPersistedApplicationDefaultEndpoints();
}
[ComImport]
[Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
public class PolicyConfigClient
{
}