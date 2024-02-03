namespace WACS.Core.SpreadSheets {
    static class Executor {
        static public void Run(string target) {
            var spreads = new List<SpreadSheet>() {
                new MovimentoDiario(),
                new Otica(),
                new Primeiras(),
                new Vendas()
            };

            foreach (string file in Directory.GetFiles(target, "*.*",SearchOption.AllDirectories))
            {
                if (!file.Contains("~$")) {
                    foreach (var spread in spreads) {
                        spread.Run(target, file);
                    }
                } 
            }
        }
    }
}