using Serilog;

namespace ScannerGUIv3.Services;

public class FileHandling
{
    public static List<string> GetFilesInFolder(string folderPath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                throw new ArgumentException(@"Folder path cannot be null or empty.", nameof(folderPath));
            }

            if (!Directory.Exists(folderPath))
            {
                throw new DirectoryNotFoundException($"The specified folder path does not exist: {folderPath}");
            }

            var files = Directory.GetFiles(folderPath).ToList();
            return files;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error occurred while getting files in folder: {FolderPath}", folderPath);
            return new List<string>();
        }
    }

    public static string GetMostRecentValidRosterFile(List<string> files)
    {
        try
        {
            if (files == null || !files.Any())
            {
                throw new ArgumentException("The list of files cannot be null or empty.");
            }

            // Step 1: Get the current time
            DateTime currentTime = DateTime.Now;

            // Step 2: Determine the valid roster period
            // Rosters are valid from 6am on a given day to 6am the next day
            DateTime rosterStart;
            if (currentTime.Hour < 6)
            {
                // If it's before 6am, the valid roster started at 6am the previous day
                rosterStart = currentTime.Date.AddDays(-1).AddHours(6);
            }
            else
            {
                // Otherwise, the valid roster started at 6am today
                rosterStart = currentTime.Date.AddHours(6);
            }
            DateTime rosterEnd = rosterStart.AddHours(24);

            // Step 3: Parse the file names and find the most recent valid file
            string mostRecentFile = null;
            DateTime mostRecentTime = DateTime.MinValue;

            foreach (var file in files)
            {
                // Extract the file name without the path
                string fileName = Path.GetFileName(file);
                // Extract the timestamp part (Roster_ddMMHHmm.json)
                if (!fileName.StartsWith("Roster_") || !fileName.EndsWith(".json"))
                {
                    continue; // Skip files that don't match the expected format
                }

                string timestampStr = fileName.Substring(7, 8); // e.g., "24031922"
                if (timestampStr.Length != 8)
                {
                    continue; // Skip if the timestamp part is not the correct length
                }

                // Parse the timestamp (ddMMHHmm)
                if (!int.TryParse(timestampStr.Substring(0, 2), out int day) ||
                    !int.TryParse(timestampStr.Substring(2, 2), out int month) ||
                    !int.TryParse(timestampStr.Substring(4, 2), out int hour) ||
                    !int.TryParse(timestampStr.Substring(6, 2), out int minute))
                {
                    continue; // Skip if parsing fails
                }

                // Assume the year is the current year (or adjust based on your needs)
                int year = currentTime.Year;
                // If the month is later than the current month, it might be from the previous year
                if (month > currentTime.Month)
                {
                    year--;
                }

                DateTime fileTime;
                try
                {
                    fileTime = new DateTime(year, month, day, hour, minute, 0);
                }
                catch (ArgumentOutOfRangeException)
                {
                    continue; // Skip if the date is invalid
                }

                // Step 4: Check if the file is within the valid roster period
                // A file is valid if it was created after the roster start time
                // and before the current time
                if (fileTime >= rosterStart && fileTime <= currentTime)
                {
                    // Step 5: Update the most recent file if this one is newer
                    if (fileTime > mostRecentTime)
                    {
                        mostRecentTime = fileTime;
                        mostRecentFile = file;
                    }
                }
            }

            if (mostRecentFile == null)
            {
                throw new FileNotFoundException("No valid roster file found for the current time.");
            }

            return mostRecentFile;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error occurred while finding the most recent valid roster file.");
            return null;
        }
    }
}