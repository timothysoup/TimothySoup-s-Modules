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
    /// <returns>
    /// A linear level as reported by Voicemeeter (usually 0–1, can exceed 1 on peaks).
    /// Returns 0 on error.
    /// </returns>
    private static float GetSafeLevel(Voicemeeter.VoicemeeterLevelType kind, int channelIndex)
    {
        if (Voicemeeter.GetLevel(kind, channelIndex, out float level) == 0)
            return level < 0f ? 0f : level;

        return 0f;
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
        SelectedChannelLevel,
    }

    protected void RegisterParameters()
    {
        RegisterParameter<float>(VoicemeeterOSCParameter.SelectedChannelLevel, "VRCOSC/Voicemeeter/Selected/Level", ParameterMode.Write, "Selected channel Level", "The loudness/level/volume of the selected channel");
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
        PollRate,
        InstallLocation,
        EnableSmoothing,
        AmplificationMultiplier,
        SmoothAttack,
        SmoothRelease,
        SoloMode,
        SelectedChannel
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

    enum PollRate
    {
        _1Hz,
        _2Hz,
        _5Hz,
        _10Hz,
        _20Hz,
    }

    // NOTE: order matters. Tick() uses range checks on this enum.
    // Do NOT reorder or insert in the middle.
    enum SoloChannel
    {
        None = -1, Phys1, Phys2, Phys3, Phys4, Phys5, Virt1, Virt2, Virt3, BusA1, BusA2, BusA3, BusA4, BusA5, BusB1, BusB2, BusB3
    }

    private int WAITING_TICKS = 0;
    private bool ENABLE_SMOOTHING = true;
    private float SMOOTH_ATTACK = 0.5f; // rise faster
    private float SMOOTH_RELEASE = 0.2f; // fall slower
    private float AMPLIFICATION = 1f;
    private int PHYS_STRIP_COUNT = 5;
    private int VIRT_STRIP_COUNT = 3;
    private int OUT_BUS_COUNT = 5;
    private int IN_BUS_COUNT = 3;
    private Voicemeeter.VoicemeeterLevelType LEVEL_TYPE = Voicemeeter.VoicemeeterLevelType.PostMuteInput;
    private bool SOLO_MODE = true;
    private SoloChannel SELECTED_CHANNEL = SoloChannel.Virt1;

    protected void CreateSettings()
    {
        CreateDropdown(VoicemeeterOSCSetting.VoicemeeterEdition, "Voicemeeter Edition", "The Voicemeeter version/edition/distribution you are using.", VoicemeeterEdition.Banana);
        CreateDropdown(VoicemeeterOSCSetting.InputLevelKind, "Input Level Kind", "Which level to get from the strips. Default: PostMute", InputLevelKind.PostMute);
        CreateDropdown(VoicemeeterOSCSetting.PollRate, "Poll rate", "How fast to poll voicemeeter in Hz", PollRate._20Hz);
        CreateTextBox(VoicemeeterOSCSetting.InstallLocation, "Voicemeeter Installation Path", "The directory at which Voicemeeter is installed. Default: C:\\Program Files (x86)\\VB\\Voicemeeter", "C:\\Program Files (x86)\\VB\\Voicemeeter");

        CreateToggle(VoicemeeterOSCSetting.SoloMode, "Solo Mode", "Only listen to selected strip/bus.", true);
        CreateDropdown(VoicemeeterOSCSetting.SelectedChannel, "Selected Channel", "Strip/Bus for the \"Selected\" parameter.", SoloChannel.Virt1);


        CreateSlider(VoicemeeterOSCSetting.AmplificationMultiplier, "Amplification", "Amplify the output levels Default: 1", 1f, 0f, 50f);

        CreateToggle(VoicemeeterOSCSetting.EnableSmoothing, "Enable Smoothing", "", true);
        CreateSlider(VoicemeeterOSCSetting.SmoothAttack, "Smooth Attack", "How fast the level should go up when volume increases. Default: 0.5", 0.5f, 0f, 1f);
        CreateSlider(VoicemeeterOSCSetting.SmoothRelease, "Smooth Release", "How fast the level should go down when volume decreases. Default: 0.2", 0.2f, 0f, 1f);
    }

    protected void CreateSettingsGroups()
    {
        CreateGroup("Voicemeeter Settings", VoicemeeterOSCSetting.VoicemeeterEdition, VoicemeeterOSCSetting.InputLevelKind, VoicemeeterOSCSetting.PollRate, VoicemeeterOSCSetting.InstallLocation);
        CreateGroup("Solo Channel Settings", VoicemeeterOSCSetting.SoloMode, VoicemeeterOSCSetting.SelectedChannel);
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
        string path = GetSettingValue<string>(VoicemeeterOSCSetting.InstallLocation) ?? "C:\\Program Files (x86)\\VB\\Voicemeeter";

        if (!string.IsNullOrWhiteSpace(path) && !SetDllDirectory(path)) // Add search path for VoicemeeterRemote64.dll
            Log($"Warning: could not set DLL directory to {path}. Is the install location correct?");

        var r = Voicemeeter.Login();
        if (r != 0) { Log($"Voicemeeter login failed: {r}"); return Task.FromResult(false); }
        Log("Voicemeeter login OK.");
        ApplySettings();
        return Task.FromResult(true);
    }

    protected void ApplySettings()
    {
        WAITING_TICKS = GetWaitingTicks();
        ENABLE_SMOOTHING = GetSettingValue<bool>(VoicemeeterOSCSetting.EnableSmoothing);
        SMOOTH_ATTACK = GetSettingValue<float>(VoicemeeterOSCSetting.SmoothAttack);
        SMOOTH_RELEASE = GetSettingValue<float>(VoicemeeterOSCSetting.SmoothRelease);
        AMPLIFICATION = GetSettingValue<float>(VoicemeeterOSCSetting.AmplificationMultiplier);
        PHYS_STRIP_COUNT = GetPhysStripCount();
        VIRT_STRIP_COUNT = GetVirtStripCount();
        OUT_BUS_COUNT = GetOutBusCount();
        IN_BUS_COUNT = GetInBusCount();
        LEVEL_TYPE = GetSettingValue<InputLevelKind>(VoicemeeterOSCSetting.InputLevelKind) switch
        {
            InputLevelKind.PreFader => Voicemeeter.VoicemeeterLevelType.PreFaderInput,
            InputLevelKind.PostFader => Voicemeeter.VoicemeeterLevelType.PostFaderInput,
            InputLevelKind.PostMute => Voicemeeter.VoicemeeterLevelType.PostMuteInput,
            _ => Voicemeeter.VoicemeeterLevelType.PostMuteInput
        };
        SOLO_MODE = GetSettingValue<bool>(VoicemeeterOSCSetting.SoloMode);
        SELECTED_CHANNEL = GetSettingValue<SoloChannel>(VoicemeeterOSCSetting.SelectedChannel);
    }

    protected override Task OnModuleStop()
    {
        try { _ = Voicemeeter.Logout(); } catch { /* ignore */ }
        return Task.CompletedTask;
    }

    // Program ticks every 50 ms
    // WAITING_TICKS = (50ms * (N+1)) roughly
    protected int GetWaitingTicks()
    {
        switch (GetSettingValue<PollRate>(VoicemeeterOSCSetting.PollRate))
        {
            case PollRate._20Hz: return 0; // 50ms
            case PollRate._10Hz: return 1; // 100ms
            case PollRate._5Hz: return 3;  // 200ms
            case PollRate._2Hz: return 9;  // 500ms
            case PollRate._1Hz: return 19; // 1000ms
            default: return 0;
        }
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
    /// Returns the Channel index for a Speaker Bus
    /// </summary>
    /// <param name="n">Bus index</param>
    private int GetOutBus(int n)
    {
        return n * 8;
    }

    /// <summary>
    /// Returns the Channel index for a Mic Bus
    /// </summary>
    /// <param name="n">Bus index</param>
    private int GetInBus(int n)
    {
        return OUT_BUS_COUNT * 8 + (n * 8);
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
                return 0;
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
                return 0;
        }
    }

    /// <summary>
    /// Returns the amount of Output (Speaker) Buses in this Voicemeeter Edition
    /// </summary>
    private int GetOutBusCount()
    {
        switch (GetSettingValue<VoicemeeterEdition>(VoicemeeterOSCSetting.VoicemeeterEdition))
        {
            case VoicemeeterEdition.Standard:
                return 1;
            case VoicemeeterEdition.Banana:
                return 3;
            case VoicemeeterEdition.Potato:
                return 5;
            default:
                Log($"GetOutBusCount Error: Unexpected Voicemeeter Edition");
                return 0;
        }
    }

    /// <summary>
    /// Returns the amount of Input (Microphone) Buses in this Voicemeeter Edition
    /// </summary>
    private int GetInBusCount()
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
                Log($"GetInBusCount Error: Unexpected Voicemeeter Edition");
                return 0;
        }
    }

    #endregion

    #region Voicemeeter Param Update (Main Logic)
    private const int MAX_PHYS = 5;
    private const int MAX_VIRT = 3;
    private const int MAX_OUT = 5;
    private const int MAX_IN = 3;

    // NOTE: we use MAX_* here (not the edition-specific counts) so parameter enums stay stable.
    private const int PHYS_INDEX_OFFSET = 0;
    private const int VIRT_INDEX_OFFSET = PHYS_INDEX_OFFSET + MAX_PHYS;
    private const int OUT_BUS_INDEX_OFFSET = VIRT_INDEX_OFFSET + MAX_VIRT;
    private const int IN_BUS_INDEX_OFFSET = OUT_BUS_INDEX_OFFSET + MAX_OUT;


    private readonly float[] _physSmooth = new float[MAX_PHYS];
    private readonly float[] _virtSmooth = new float[MAX_VIRT];
    private readonly float[] _outBusSmooth = new float[MAX_OUT];
    private readonly float[] _inBusSmooth = new float[MAX_IN];

    private const float SMOOTH_FLOOR = 0.00005f;

    /// <summary>
    /// Returns a smoothed float value based on its previous value
    /// </summary>
    /// <param name="prev">The previous value</param>
    /// <param name="input">The current value</param>
    private float Smooth(float prev, float input)
    {
        if (!ENABLE_SMOOTHING) return input;
        float alpha = input > prev ? SMOOTH_ATTACK : SMOOTH_RELEASE;
        float res = prev + alpha * (input - prev);
        return res < SMOOTH_FLOOR ? 0 : res;
    }

    private const float MAX_LEVEL_OUTPUT = 10f;
    private const float MIN_LEVEL_OUTPUT = 0f;

    /// <summary>
    /// Updates a parameter, synchronizing it to the Voicemeeter level.
    /// Note: values are clamped to [0, 10] before being sent to VRCOSC.
    /// </summary>
    /// <param name="index">The strip/bus index</param>
    /// <param name="chL">The channel index</param>
    /// <param name="levelType">The kind of level to get with GetLevel</param>
    /// <param name="smoothingState">The smoothing array used</param>
    /// <param name="solo">Whether or not to push this update to the Solo parameter</param>
    private void UpdateParameter(int index, int chL, int offset, Voicemeeter.VoicemeeterLevelType levelType, float[] smoothingState, bool solo = false)
    {
        int chR = chL + 1;

        float l = GetSafeLevel(levelType, chL);
        float r = GetSafeLevel(levelType, chR);

        float raw = MathF.Max(l, r) * AMPLIFICATION;

        smoothingState[index] = Smooth(smoothingState[index], raw);

        float clamped = Math.Clamp(smoothingState[index], MIN_LEVEL_OUTPUT, MAX_LEVEL_OUTPUT);

        if (solo) SendParameter(VoicemeeterOSCParameter.SelectedChannelLevel, clamped);
        SendParameter((VoicemeeterOSCParameter)(index + offset), clamped);
    }

    private int _waitedTicks = 0;

    // poll every 50ms
    [ModuleUpdate(ModuleUpdateMode.Custom, updateImmediately: true, deltaMilliseconds: 50)]
    protected void Tick()
    {
        if (_waitedTicks < WAITING_TICKS)
        {
            _waitedTicks++;
            return;
        }
        _waitedTicks = 0;
        if (SOLO_MODE)
        {
            int s = (int)SELECTED_CHANNEL;
            if (s < 0) return;
            if (SELECTED_CHANNEL < SoloChannel.Virt1)
            {
                // It is a physical strip
                UpdateParameter(s, GetPhysStrip(s), PHYS_INDEX_OFFSET, LEVEL_TYPE, _physSmooth, true);
            }
            else if (SELECTED_CHANNEL < SoloChannel.BusA1)
            {
                // It is a virtual strip
                // Reset to Virtual Strip index 0
                s -= (int)SoloChannel.Virt1;
                UpdateParameter(s, GetVirtStrip(s), VIRT_INDEX_OFFSET, LEVEL_TYPE, _virtSmooth, true);
            }
            else if (SELECTED_CHANNEL < SoloChannel.BusB1)
            {
                // It is a speaker output bus
                // Reset to Speaker Bus index 0
                s -= (int)SoloChannel.BusA1;
                UpdateParameter(s, GetOutBus(s), OUT_BUS_INDEX_OFFSET, Voicemeeter.VoicemeeterLevelType.OutputBus, _outBusSmooth, true);
            }
            else
            {
                // It is a mic output bus
                // Reset to Mic Bus index 0
                s -= (int)SoloChannel.BusB1;
                UpdateParameter(s, GetInBus(s), IN_BUS_INDEX_OFFSET, Voicemeeter.VoicemeeterLevelType.OutputBus, _inBusSmooth, true);
            }
        }
        else
        {
            int o = 0;
            // --- STRIPS (Physical inputs) ---
            for (int s = 0; s < PHYS_STRIP_COUNT; s++)
            {
                UpdateParameter(s, GetPhysStrip(s), PHYS_INDEX_OFFSET, LEVEL_TYPE, _physSmooth, (int)SELECTED_CHANNEL == s + o);
            }
            o += PHYS_STRIP_COUNT;

            // --- STRIPS (Virtual inputs) ---
            for (int s = 0; s < VIRT_STRIP_COUNT; s++)
            {
                UpdateParameter(s, GetVirtStrip(s), VIRT_INDEX_OFFSET, LEVEL_TYPE, _virtSmooth, (int)SELECTED_CHANNEL == s + o);
            }
            o += VIRT_STRIP_COUNT;

            // --- Speaker BUSES (outputs) ---
            for (int b = 0; b < OUT_BUS_COUNT; b++)
            {
                UpdateParameter(b, GetOutBus(b), OUT_BUS_INDEX_OFFSET, Voicemeeter.VoicemeeterLevelType.OutputBus, _outBusSmooth, (int)SELECTED_CHANNEL == b + o);
            }
            o += OUT_BUS_COUNT;

            // --- Mic BUSES (outputs) ---
            for (int b = 0; b < IN_BUS_COUNT; b++)
            {
                UpdateParameter(b, GetInBus(b), IN_BUS_INDEX_OFFSET, Voicemeeter.VoicemeeterLevelType.OutputBus, _inBusSmooth, (int)SELECTED_CHANNEL == b + o);
            }
        }
    }

    #endregion
}
