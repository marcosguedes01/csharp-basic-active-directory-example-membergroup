using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;

namespace ADTest
{
    /// <summary>
    /// Fonte:
    ///     https://social.technet.microsoft.com/wiki/contents/articles/5392.active-directory-ldap-syntax-filters.aspx
    /// </summary>
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            Execute();
        }

        private static void Execute()
        {
            try
            {
                List<ADGroup> lstResult = new List<ADGroup>();

                string domainUrl = "";
                string domainOU = "";
                string domainDC = "";
                string username = "";
                string password = "";

                string samAccountName = "";
                int maxCount = 5000;

                using (DirectoryEntry de = new DirectoryEntry("LDAP://" + domainUrl + domainOU + domainDC, username, password, AuthenticationTypes.None))
                {
                    DirectorySearcher search = new DirectorySearcher(de);
                    search.PageSize = maxCount;

                    string query = "(sAMAccountName=" + samAccountName + ")";
                    search.Filter = query;
                    search.PropertiesToLoad.Add("memberOf");

                    // Pegar os grupos do usuário consultado
                    SearchResult searchResult = search.FindOne();

                    if (searchResult != null && searchResult.Properties.Contains("memberOf"))
                    {
                        foreach (string prop in searchResult.Properties["memberOf"])
                        {
                            ADGroup _ad = new ADGroup();

                            string cnGroup = string.Empty;
                            int indexSeparator = prop.IndexOf(",");

                            if (indexSeparator > 0)
                            {
                                cnGroup = prop.Substring(0, indexSeparator);
                            }
                            else
                            {
                                cnGroup = prop.Substring(0);
                            }

                            DirectorySearcher searchG = new DirectorySearcher(de);
                            searchG.PageSize = maxCount;

                            query = "(&(objectCategory=group)(" + cnGroup + "))";
                            searchG.Filter = query;
                            searchG.PropertiesToLoad.Add("name");
                            searchG.PropertiesToLoad.Add("description");
                            searchG.PropertiesToLoad.Add("distinguishedName");

                            // Obter detalhes do grupo
                            SearchResult searchGResult = searchG.FindOne();

                            if (searchGResult != null && searchGResult.Properties.Contains("name"))
                            {
                                _ad.Name = searchGResult.Properties["name"][0].ToString();
                            }

                            if (searchGResult != null && searchGResult.Properties.Contains("description"))
                            {
                                _ad.Description = searchGResult.Properties["description"][0].ToString();
                            }

                            if (searchGResult != null && searchGResult.Properties.Contains("distinguishedName"))
                            {
                                _ad.DistinguishedName = searchGResult.Properties["distinguishedName"][0].ToString();
                            }

                            DirectorySearcher searchU = new DirectorySearcher(de);
                            searchU.PageSize = maxCount;

                            query = "(memberOf=" + _ad.DistinguishedName + ")";
                            searchU.Filter = query;
                            searchU.PropertiesToLoad.Add("sAMAccountName");
                            searchU.PropertiesToLoad.Add("displayName");

                            SearchResultCollection usersResult = searchU.FindAll();

                            if (usersResult != null)
                            {
                                foreach (SearchResult _sResult in usersResult)
                                {
                                    User _user = new User();

                                    if (_sResult.Properties.Contains("sAMAccountName"))
                                    {
                                        _user.SAMAccountName = _sResult.Properties["sAMAccountName"][0].ToString();
                                    }

                                    if (_sResult.Properties.Contains("displayName"))
                                    {
                                        _user.DisplayName = _sResult.Properties["displayName"][0].ToString();
                                    }

                                    _ad.Users.Add(_user);
                                }
                            }

                            _ad.Users = _ad.Users.OrderBy(u => u.DisplayName).ToList();
                            lstResult.Add(_ad);
                        }
                    }
                }

                lstResult = lstResult.OrderBy(r => r.Name).ToList();
            }
            catch (Exception ex) { }
        }
    }

    public class ADGroup
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string DistinguishedName { get; set; }
        public List<User> Users { get; set; }

        public ADGroup()
        {
            this.Users = new List<User>();
        }

        public override string ToString()
        {
            return Name ?? base.ToString();
        }
    }

    public class User
    {
        public string SAMAccountName { get; set; }
        public string DisplayName { get; set; }

        public override string ToString()
        {
            return string.Concat(SAMAccountName, " - ", DisplayName);
        }
    }

}
