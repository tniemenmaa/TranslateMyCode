using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslateMyCode.Translate
{
    public interface ITranslator
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="word"></param>
        /// <returns></returns>
        string Translate(string word, Language from, Language to);
    }
}
