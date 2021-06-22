using System;
using System.Collections.Generic;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Enums;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    class DocumentGeneratorsFactory
    {
        private static Dictionary<DocumentGeneratorType, Type> generatorTable => new Dictionary<DocumentGeneratorType, Type>
        {
            {DocumentGeneratorType.XmlDocument, typeof(XmlDocumentGenerator)}
        };
        private readonly DocumentGeneratorType generatorType;
        public DocumentGeneratorsFactory(DocumentGeneratorType generatorType)
        {
            this.generatorType = generatorType;
        }
        public IDocumentGenerator<T> GetDocumentGenerator<T>()
        {
            if (!generatorTable.ContainsKey(this.generatorType)) throw new NotSupportedException("this type of documents isn't supported");
            var generatorClass = generatorTable[this.generatorType];
            IDocumentGenerator<T> generatorInstance = (IDocumentGenerator<T>) Activator.CreateInstance(generatorClass);
            return generatorInstance;
        }
    }
}
