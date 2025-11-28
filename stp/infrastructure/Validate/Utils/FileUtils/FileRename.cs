namespace infrastructure.Validate.Utils.FileUtils
{
    class FileRename
    {
        public static String Rename(String filePath, String name)
        {
            if (!File.Exists(filePath))
                return filePath;

            String directory = Path.GetDirectoryName(filePath)!;
            String newPath = Path.Combine(directory, name);

            try
            {
                File.Move(filePath, newPath, overwrite: true);
                return newPath;
            }
            catch (Exception)
            {
                return filePath;
            }
        }
    }
}