using System.Runtime.InteropServices;

namespace TimothySoup.Modules.VoicemeeterOSC;

internal static class Voicemeeter
{
    public enum VoicemeeterLevelType
    {
        /// <summary>
        /// Input strips, pre-fader & pre-mute (raw input signal)
        /// </summary>
        PreFaderInput = 0,

        /// <summary>
        /// Input strips, post-fader (affected by fader, before mute)
        /// </summary>
        PostFaderInput = 1,

        /// <summary>
        /// Input strips, post-mute (signal after mute state)
        /// </summary>
        PostMuteInput = 2,

        /// <summary>
        /// Output buses (post-fader / post-mute)
        /// </summary>
        OutputBus = 3
    }

    private const string Dll = "VoicemeeterRemote64.dll";

    [DllImport(Dll, EntryPoint = "VBVMR_Login", CallingConvention = CallingConvention.StdCall)]
    public static extern int Login();

    [DllImport(Dll, EntryPoint = "VBVMR_Logout", CallingConvention = CallingConvention.StdCall)]
    public static extern int Logout();

    [DllImport(Dll, EntryPoint = "VBVMR_GetLevel", CallingConvention = CallingConvention.StdCall)]
    public static extern int GetLevel(VoicemeeterLevelType kind, int channelIndex, out float level);
}
