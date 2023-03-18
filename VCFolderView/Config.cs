using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VCFolderView
{
    public class Config
    {
        public string WorkingDirectory;
        public List<Regex> ExcludeRules;

        public bool IsExcluded(string path)
        {
            foreach (Regex regex in ExcludeRules)
            {
                if (regex.IsMatch(path))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
