using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Installer.Shared
{
    public static class LanguageCodeConverter
    {
        // Cache of cultures to avoid repeated calls to GetCultures
        private static readonly CultureInfo[] AllCultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
        
        // Get language code from language name (e.g., "English" -> "en")
        public static string GetLanguageCodeFromName(string languageName)
        {
            if (string.IsNullOrWhiteSpace(languageName))
                return null;

            try
            {
                // First try exact match on DisplayName or EnglishName
                var culture = AllCultures.FirstOrDefault(c => 
                    c.DisplayName.Equals(languageName, StringComparison.OrdinalIgnoreCase) ||
                    c.EnglishName.Equals(languageName, StringComparison.OrdinalIgnoreCase) ||
                    c.NativeName.Equals(languageName, StringComparison.OrdinalIgnoreCase));
                
                if (culture != null)
                    return culture.TwoLetterISOLanguageName;
                
                // Then try contains match
                culture = AllCultures.FirstOrDefault(c => 
                    c.DisplayName.StartsWith(languageName, StringComparison.OrdinalIgnoreCase) ||
                    c.EnglishName.StartsWith(languageName, StringComparison.OrdinalIgnoreCase) ||
                    c.NativeName.StartsWith(languageName, StringComparison.OrdinalIgnoreCase));
                
                if (culture != null)
                    return culture.TwoLetterISOLanguageName;
                
                // Try more flexible contains match
                culture = AllCultures.FirstOrDefault(c => 
                    c.DisplayName.Contains(languageName, StringComparison.OrdinalIgnoreCase) ||
                    c.EnglishName.Contains(languageName, StringComparison.OrdinalIgnoreCase) ||
                    c.NativeName.Contains(languageName, StringComparison.OrdinalIgnoreCase));
                
                if (culture != null)
                    return culture.TwoLetterISOLanguageName;
            }
            catch
            {
                // Ignore errors and return null
            }
            
            return null;
        }

        // Check if a locale code matches a language name (e.g., "en-US" matches "English")
        public static bool LocaleMatchesLanguage(string locale, string languageName)
        {
            if (string.IsNullOrWhiteSpace(locale) || string.IsNullOrWhiteSpace(languageName))
                return false;
                
            // Direct match on locale
            if (locale.Equals(languageName, StringComparison.OrdinalIgnoreCase))
                return true;
                
            // Try to get language code from name
            string languageCode = GetLanguageCodeFromName(languageName);
            if (string.IsNullOrWhiteSpace(languageCode))
                return false;
                
            // Check if locale starts with the language code
            return locale.StartsWith(languageCode + "-", StringComparison.OrdinalIgnoreCase);
        }

        public static int ConvertToLcid(string langCode, string country)
        {
            // Placeholder logic; replace with actual LCID mapping
            return 1033; // Default to en-US LCID
        }
    }
}
