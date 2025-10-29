using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using VRCOSC.App.SDK.Modules;
using VRCOSC.App.SDK.Parameters;

namespace TimothySoup.Modules.VoicemeeterOSC;

[ModuleTitle("Voicemeeter Levels")]
[ModuleDescription("Reads volume/levels from Voicemeeter channels")]
[ModuleType(ModuleType.Generic)]
public class VoicemeeterOSC : Module
{
    #region __internal__
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool SetDllDirectory(string lpPathName);

    /// <summary>
    /// Reads the Level from Voicemeeter using GetLevel
    /// </summary>
    /// <param name="kind">Where to read the level from</param>
    /// <param name="channelIndex">Which channel to read it from</param>
    /// <returns>A normalized dB float</returns>
    private static float ReadDb(Voicemeeter.VoicemeeterLevelType kind, int channelIndex)
    {
        // Returns dB or -200 on error (treat as silence)
        return Voicemeeter.GetLevel(kind, channelIndex, out float valDb) == 0 ? valDb : -200f;
    }
    #endregion

    #region Parameters
    enum VoicemeeterOSCParameter
    {
        Strip0Level,
        Strip1Level,
        Strip2Level,
        Strip3Level,
        Strip4Level,
        Strip5Level,
        Strip6Level,
        Strip7Level,
        Bus0Level,
        Bus1Level,
        Bus2Level,
        Bus3Level,
        Bus4Level,
        Bus5Level,
        Bus6Level,
        Bus7Level,
        SoloChannel,
    }

    protected void RegisterParameters()
    {
        RegisterParameter<float>(VoicemeeterOSCParameter.SoloChannel, "VRCOSC/Voicemeeter/SoloChannel/Level", ParameterMode.Write, "Solo selected channel", "The loudness/level/volume of the selected Solo channel");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip0Level, "VRCOSC/Voicemeeter/Phys/1/Level", ParameterMode.Write, "Physical Strip 1 Level", "The loudness/level/volume of Physical Strip 1 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip1Level, "VRCOSC/Voicemeeter/Phys/2/Level", ParameterMode.Write, "Physical Strip 2 Level", "The loudness/level/volume of Physical Strip 2 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip2Level, "VRCOSC/Voicemeeter/Phys/3/Level", ParameterMode.Write, "Physical Strip 3 Level", "The loudness/level/volume of Physical Strip 3 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip3Level, "VRCOSC/Voicemeeter/Phys/4/Level", ParameterMode.Write, "Physical Strip 4 Level", "The loudness/level/volume of Physical Strip 4 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip4Level, "VRCOSC/Voicemeeter/Phys/5/Level", ParameterMode.Write, "Physical Strip 5 Level", "The loudness/level/volume of Physical Strip 5 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip5Level, "VRCOSC/Voicemeeter/Virt/1/Level", ParameterMode.Write, "Virtual Strip 1 Level", "The loudness/level/volume of Virtual Strip 1 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip6Level, "VRCOSC/Voicemeeter/Virt/2/Level", ParameterMode.Write, "Virtual Strip 2 Level", "The loudness/level/volume of Virtual Strip 2 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Strip7Level, "VRCOSC/Voicemeeter/Virt/3/Level", ParameterMode.Write, "Virtual Strip 3 Level", "The loudness/level/volume of Virtual Strip 3 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus0Level, "VRCOSC/Voicemeeter/Bus/A1/Level", ParameterMode.Write, "Bus A1 Level", "The loudness/level/volume of BUS A1 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus1Level, "VRCOSC/Voicemeeter/Bus/A2/Level", ParameterMode.Write, "Bus A2 Level", "The loudness/level/volume of BUS A2 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus2Level, "VRCOSC/Voicemeeter/Bus/A3/Level", ParameterMode.Write, "Bus A3 Level", "The loudness/level/volume of BUS A3 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus3Level, "VRCOSC/Voicemeeter/Bus/A4/Level", ParameterMode.Write, "Bus A4 Level", "The loudness/level/volume of BUS A4 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus4Level, "VRCOSC/Voicemeeter/Bus/A5/Level", ParameterMode.Write, "Bus A5 Level", "The loudness/level/volume of BUS A5 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus5Level, "VRCOSC/Voicemeeter/Bus/B1/Level", ParameterMode.Write, "Bus B1 Level", "The loudness/level/volume of BUS B1 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus6Level, "VRCOSC/Voicemeeter/Bus/B2/Level", ParameterMode.Write, "Bus B2 Level", "The loudness/level/volume of BUS B2 in Voicemeeter");
        RegisterParameter<float>(VoicemeeterOSCParameter.Bus7Level, "VRCOSC/Voicemeeter/Bus/B3/Level", ParameterMode.Write, "Bus B3 Level", "The loudness/level/volume of BUS B3 in Voicemeeter");
    }
    #endregion

    #region Settings

    enum VoicemeeterOSCSetting
    {
        VoicemeeterEdition,
        InputLevelKind,
        EnableSmoothing,
        AmplificationMultiplier,
        SmoothAttack,
        SmoothRelease,
        SoloMode,
        SoloChannel
    }

    enum VoicemeeterEdition
    {
        Standard,
        Banana,
        Potato
    }

    enum InputLevelKind
    {
        PreFader = 0,
        PostFader = 1,
        PostMute = 2
    }

    enum SoloChannel
    {
        Phys1, Phys2, Phys3, Phys4, Phys5, Virt1, Virt2, Virt3, BusA1, BusA2, BusA3, BusA4, BusA5, BusB1, BusB2, BusB3
    }

    private bool ENABLE_SMOOTHING = true;
    private float SMOOTH_ATTACK = 0.5f; // rise faster
    private float SMOOTH_RELEASE = 0.2f; // fall slower
    private float AMPLIFICATION = 1f;
    private int PHYS_STRIP_COUNT = 5;
    private int VIRT_STRIP_COUNT = 3;
    private int BUS_COUNT = 8;
    private Voicemeeter.VoicemeeterLevelType LEVEL_TYPE = Voicemeeter.VoicemeeterLevelType.PostMuteInput;
    private bool SOLO_MODE = true;
    private SoloChannel SOLO_CHANNEL = SoloChannel.Virt1;

    protected void CreateSettings()
    {
        CreateDropdown(VoicemeeterOSCSetting.VoicemeeterEdition, "Voicemeeter Edition", "The Voicemeeter version/edition/distribution you are using.", VoicemeeterEdition.Banana);
        CreateDropdown(VoicemeeterOSCSetting.InputLevelKind, "Input Level Kind", "Which level to get from the strips. Default: PostMute", InputLevelKind.PostMute);

        CreateToggle(VoicemeeterOSCSetting.SoloMode, "Solo Mode", "Only listen to one strip/bus.", true);
        CreateDropdown(VoicemeeterOSCSetting.SoloChannel, "Solo Channel", "Strip/Bus to listen to in Solo Mode.", SoloChannel.Virt1);


        CreateSlider(VoicemeeterOSCSetting.AmplificationMultiplier, "Amplification", "Amplify the output levels Default: 1", 1f, 0f, 50f);

        CreateToggle(VoicemeeterOSCSetting.EnableSmoothing, "Enable Smoothing", "", true);
        CreateSlider(VoicemeeterOSCSetting.SmoothAttack, "Smooth Attack", "How fast the level should go up when volume increases. Default: 0.5", 0.5f, 0f, 1f);
        CreateSlider(VoicemeeterOSCSetting.SmoothRelease, "Smooth Release", "How fast the level should go down when volume decreases. Default: 0.2", 0.2f, 0f, 1f);
    }

    protected void CreateSettingsGroups()
    {
        CreateGroup("Voicemeeter Settings", VoicemeeterOSCSetting.VoicemeeterEdition, VoicemeeterOSCSetting.InputLevelKind);
        CreateGroup("Solo Channel Settings", VoicemeeterOSCSetting.SoloMode, VoicemeeterOSCSetting.SoloChannel);
        CreateGroup("Level Settings", VoicemeeterOSCSetting.AmplificationMultiplier, VoicemeeterOSCSetting.EnableSmoothing, VoicemeeterOSCSetting.SmoothAttack, VoicemeeterOSCSetting.SmoothRelease);
    }

    #endregion

    #region Module Config
    protected override void OnPreLoad()
    {
        RegisterParameters();
        CreateSettings();
        CreateSettingsGroups();
    }

    protected override Task<bool> OnModuleStart()
    {
        ApplySettings();
        SetDllDirectory("C:\\Program Files (x86)\\VB\\Voicemeeter");  // <- add search path for VoicemeeterRemote64.dll

        var r = Voicemeeter.Login();
        if (r != 0) { Log($"Voicemeeter login failed: {r}"); return Task.FromResult(false); }
        Log("Voicemeeter login OK.");
        return Task.FromResult(true);
    }

    protected void ApplySettings()
    {
        ENABLE_SMOOTHING = GetSettingValue<bool>(VoicemeeterOSCSetting.EnableSmoothing);
        SMOOTH_ATTACK = GetSettingValue<float>(VoicemeeterOSCSetting.SmoothAttack);
        SMOOTH_RELEASE = GetSettingValue<float>(VoicemeeterOSCSetting.SmoothRelease);
        AMPLIFICATION = GetSettingValue<float>(VoicemeeterOSCSetting.AmplificationMultiplier);
        PHYS_STRIP_COUNT = GetPhysStripCount();
        VIRT_STRIP_COUNT = GetVirtStripCount();
        BUS_COUNT = GetBusCount();
        LEVEL_TYPE = (Voicemeeter.VoicemeeterLevelType)InputLevelKind.PostMute;
        SOLO_MODE = GetSettingValue<bool>(VoicemeeterOSCSetting.SoloMode);
        SOLO_CHANNEL = GetSettingValue<SoloChannel>(VoicemeeterOSCSetting.SoloChannel);
    }

    protected override Task OnModuleStop()
    {
        try { _ = Voicemeeter.Logout(); } catch { /* ignore */ }
        return Task.CompletedTask;
    }
    #endregion

    #region Compatibility Helpers
    /// <summary>
    /// Returns the Channel index for a Physical Strip
    /// </summary>
    /// <param name="n">Physical Strip index</param>
    private int GetPhysStrip(int n)
    {
        return n * 2;
    }

    /// <summary>
    /// Returns the Channel index for a Virtual Strip
    /// </summary>
    /// <param name="n">Virtual Strip index</param>
    private int GetVirtStrip(int n)
    {
        return PHYS_STRIP_COUNT * 2 + n * 8;
    }

    /// <summary>
    /// Returns the Channel index for a Bus
    /// </summary>
    /// <param name="n">Bus index</param>
    private int GetBus(int n)
    {
        return n * 8;
    }

    /// <summary>
    /// Returns the amount of Physical Strips in this Voicemeeter Edition
    /// </summary>
    private int GetPhysStripCount()
    {
        switch (GetSettingValue<VoicemeeterEdition>(VoicemeeterOSCSetting.VoicemeeterEdition))
        {
            case VoicemeeterEdition.Standard:
                return 2;
            case VoicemeeterEdition.Banana:
                return 3;
            case VoicemeeterEdition.Potato:
                return 5;
            default:
                Log($"GetPhysStripCount Error: Unexpected Voicemeeter Edition");
                return -1;
        }
    }

    /// <summary>
    /// Returns the amount of Virtual Strips in this Voicemeeter Edition
    /// </summary>
    private int GetVirtStripCount()
    {
        switch (GetSettingValue<VoicemeeterEdition>(VoicemeeterOSCSetting.VoicemeeterEdition))
        {
            case VoicemeeterEdition.Standard:
                return 1;
            case VoicemeeterEdition.Banana:
                return 2;
            case VoicemeeterEdition.Potato:
                return 3;
            default:
                Log($"GetVirtStripCount Error: Unexpected Voicemeeter Edition");
                return -1;
        }
    }

    /// <summary>
    /// Returns the amount of Output Buses in this Voicemeeter Edition
    /// </summary>
    private int GetBusCount()
    {
        switch (GetSettingValue<VoicemeeterEdition>(VoicemeeterOSCSetting.VoicemeeterEdition))
        {
            case VoicemeeterEdition.Standard:
                return 2;
            case VoicemeeterEdition.Banana:
                return 5;
            case VoicemeeterEdition.Potato:
                return 8;
            default:
                Log($"GetBusCount Error: Unexpected Voicemeeter Edition");
                return -1;
        }
    }

    #endregion

    #region Voicemeeter Param Update (Main Logic)
    private const int PHYS_INDEX_OFFSET = 0;
    private const int VIRT_INDEX_OFFSET = PHYS_INDEX_OFFSET + 5;
    private const int BUS_INDEX_OFFSET = VIRT_INDEX_OFFSET + 3;

    private readonly float[] _physSmooth = new float[5];
    private readonly float[] _virtSmooth = new float[3];
    private readonly float[] _busSmooth = new float[8];

    /// <summary>
    /// Returns a smoothened float value based on it's previous value
    /// </summary>
    /// <param name="prev">The previous value</param>
    /// <param name="input">The current value</param>
    private float Smooth(float prev, float input)
    {
        if (!ENABLE_SMOOTHING) return input;
        float alpha = input > prev ? SMOOTH_ATTACK : SMOOTH_RELEASE;
        float res = prev + alpha * (input - prev);
        return res < 0.00005 ? 0 : res;
    }

    /// <summary>
    /// Updates a parameter, synchronizing it to the Voicemeeter lever
    /// </summary>
    /// <param name="index">The strip/bus index</param>
    /// <param name="chL">The channel index</param>
    /// <param name="levelType">The kind of level to get with GetLevel</param>
    /// <param name="smoothingState">The smoothing array used</param>
    /// <param name="solo">Whether or not to push this update to the Solo parameter</param>
    private void UpdateParameter(int index, int chL, int offset, Voicemeeter.VoicemeeterLevelType levelType, float[] smoothingState, bool solo = false)
    {
        int chR = chL + 1;

        float lDb = ReadDb(levelType, chL);
        float rDb = ReadDb(levelType, chR);

        float raw = Math.Clamp(MathF.Max(lDb, rDb) * AMPLIFICATION, 0f, 1f);
        smoothingState[index] = Smooth(smoothingState[index], raw);

        if (solo) SendParameter(VoicemeeterOSCParameter.SoloChannel, smoothingState[index]);
        SendParameter((VoicemeeterOSCParameter)(index + offset), smoothingState[index]);
    }

    // poll every 50ms
    [ModuleUpdate(ModuleUpdateMode.Custom, updateImmediately: true, deltaMilliseconds: 50)]
    protected void Tick()
    {
        if (SOLO_MODE)
        {
            int s = (int)SOLO_CHANNEL;
            if (SOLO_CHANNEL < SoloChannel.Virt1)
            {
                // It is a physical strip
                UpdateParameter(s, GetPhysStrip(s), PHYS_INDEX_OFFSET, LEVEL_TYPE, _physSmooth, true);
            }
            else if (SOLO_CHANNEL < SoloChannel.BusA1)
            {
                // It is a virtual strip
                // Reset to Virtual Strip index 0
                s -= (int)SoloChannel.Virt1;
                UpdateParameter(s, GetVirtStrip(s), VIRT_INDEX_OFFSET, LEVEL_TYPE, _virtSmooth, true);
            }
            else
            {
                // It is an output bus
                // Reset to Bus index 0
                s -= (int)SoloChannel.BusA1;
                UpdateParameter(s, GetBus(s), BUS_INDEX_OFFSET, Voicemeeter.VoicemeeterLevelType.OutputBus, _busSmooth, true);
            }
        }
        else
        {
            // --- STRIPS (Phyiscal inputs) ---
            for (int s = 0; s < PHYS_STRIP_COUNT; s++)
            {
                UpdateParameter(s, GetPhysStrip(s), PHYS_INDEX_OFFSET, LEVEL_TYPE, _physSmooth);
            }

            // --- STRIPS (Virtual inputs) ---
            for (int s = 0; s < VIRT_STRIP_COUNT; s++)
            {
                UpdateParameter(s, GetVirtStrip(s), VIRT_INDEX_OFFSET, LEVEL_TYPE, _virtSmooth);
            }

            // --- BUSES (outputs) ---
            for (int b = 0; b < BUS_COUNT; b++)
            {
                UpdateParameter(b, GetBus(b), BUS_INDEX_OFFSET, Voicemeeter.VoicemeeterLevelType.OutputBus, _busSmooth);
            }
        }
    }

    #endregion
}
