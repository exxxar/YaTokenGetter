using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RegAccYandex
{
    public class User
    {
        public string name;
        public string tname;
        public string login;
        public string pass;
        public string prog_id;
        public string prog_pass;

        public string question;
        public string answer;
        public string date_reg;
        public string callback_url;

        public Dictionary<string, string> serialize()
        {
            Type myType = this.GetType();
            FieldInfo[] fields = myType.GetFields(BindingFlags.GetField | BindingFlags.Instance | BindingFlags.Public);
            Dictionary<string, string> dict = new Dictionary<string, string>();

            fields
                .ToList()
                .ForEach(f => {
                    Console.WriteLine("{0}=>{1}", f.Name, f.GetValue(this));
                    dict.Add(f.Name, Convert.ToString(f.GetValue(this)));


                    });

            return dict;
        }

        public override string ToString()
        {
            StringBuilder build = new StringBuilder();
            this.serialize()
                .ToList()
                .ForEach(s =>
                build.AppendFormat("{0};", s.Value)
            );
            return build.ToString();
        }
    }
}
