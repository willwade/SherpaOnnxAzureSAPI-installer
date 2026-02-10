using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;
using System.Threading.Tasks;

namespace SherpaOnnxConfig
{
    static class Program
    {
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);

        [DllImport("kernel32.dll")]
        private static extern bool FreeConsole();

        [STAThread]
        static int Main(string[] args)
        {
            // CLI mode: if arguments are provided, run in command-line mode
            if (args.Length > 0)
            {
                return RunCliMode(args);
            }

            // GUI mode: no arguments, show the form
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        private static int RunCliMode(string[] args)
        {
            // Try to attach to parent console first, or allocate a new one
            if (!AttachConsole(-1))
            {
                AllocConsole();
            }

            // Redirect stdout/stderr to the newly allocated console
            var stdout = Console.OpenStandardOutput();
            var stderr = Console.OpenStandardError();
            var writer = new System.IO.StreamWriter(stdout) { AutoFlush = true };
            var errorWriter = new System.IO.StreamWriter(stderr) { AutoFlush = true };
            Console.SetOut(writer);
            Console.SetError(errorWriter);

            try
            {
                var cli = new CliHandler();
                int result = cli.ExecuteAsync(args).GetAwaiter().GetResult();

                // Keep console open briefly to see output
                System.Threading.Thread.Sleep(500);
                FreeConsole();
                return result;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"ERROR: {ex.Message}");
                System.Threading.Thread.Sleep(1000);
                FreeConsole();
                return 1;
            }
        }
    }

    internal class CliHandler
    {
        public async Task<int> ExecuteAsync(string[] args)
        {
            string command = args[0].ToLower();

            switch (command)
            {
                case "--download":
                case "-d":
                    return await DownloadVoice(args);

                case "--install":
                case "-i":
                    return await InstallVoice(args);

                case "--uninstall":
                case "-u":
                    return await UninstallVoice(args);

                case "--list":
                case "-l":
                    return ListVoices(args);

                case "--help":
                case "-h":
                case "/?":
                    ShowHelp();
                    return 0;

                default:
                    Console.Error.WriteLine($"Unknown command: {command}");
                    ShowHelp();
                    return 1;
            }
        }

        private async Task<int> DownloadVoice(string[] args)
        {
            if (args.Length < 2)
            {
                Console.Error.WriteLine("ERROR: --download requires a voice ID");
                Console.Error.WriteLine("Usage: --download <voice-id>");
                return 1;
            }

            string voiceId = args[1];
            Console.WriteLine($"=== Downloading voice: {voiceId} ===");

            // This would call the same download logic as MainForm
            // For now, return not implemented
            Console.WriteLine("ERROR: CLI download not yet implemented - please use the GUI");
            return 1;
        }

        private async Task<int> InstallVoice(string[] args)
        {
            if (args.Length < 2)
            {
                Console.Error.WriteLine("ERROR: --install requires a voice ID");
                Console.Error.WriteLine("Usage: --install <voice-id>");
                return 1;
            }

            string voiceId = args[1];
            Console.WriteLine($"=== Installing voice to SAPI5: {voiceId} ===");

            // This would call the same install logic as MainForm
            Console.WriteLine("ERROR: CLI install not yet implemented - please use the GUI");
            return 1;
        }

        private async Task<int> UninstallVoice(string[] args)
        {
            if (args.Length < 2)
            {
                Console.Error.WriteLine("ERROR: --uninstall requires a voice ID");
                Console.Error.WriteLine("Usage: --uninstall <voice-id>");
                return 1;
            }

            string voiceId = args[1];
            Console.WriteLine($"=== Uninstalling voice: {voiceId} ===");

            // This would call the same uninstall logic as MainForm
            Console.WriteLine("ERROR: CLI uninstall not yet implemented - please use the GUI");
            return 1;
        }

        private int ListVoices(string[] args)
        {
            Console.WriteLine("=== Available Voices ===");
            Console.WriteLine("Run with --list --verbose for more details");
            // TODO: Load catalog and list voices
            return 0;
        }

        private void ShowHelp()
        {
            Console.WriteLine("SherpaOnnx SAPI5 Voice Installer");
            Console.WriteLine();
            Console.WriteLine("Usage: SherpaOnnxConfig.exe [command] [options]");
            Console.WriteLine();
            Console.WriteLine("Commands:");
            Console.WriteLine("  -d, --download <voice-id>    Download a voice model");
            Console.WriteLine("  -i, --install <voice-id>     Install voice to SAPI5 (requires admin)");
            Console.WriteLine("  -u, --uninstall <voice-id>   Uninstall voice from SAPI5 (requires admin)");
            Console.WriteLine("  -l, --list                   List available voices");
            Console.WriteLine("  -h, --help                   Show this help message");
            Console.WriteLine();
            Console.WriteLine("If no command is specified, the GUI will launch.");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  SherpaOnnxConfig.exe --download mms_hat");
            Console.WriteLine("  SherpaOnnxConfig.exe --install mms_hat");
            Console.WriteLine("  SherpaOnnxConfig.exe --list");
        }
    }
}
