namespace XCeedWordInspeccion
{
    public class Model
    {

        public class Ensayo
        {
           public string IdProducto { get; set; }
           public string NombreGenerico { get; set; }
           public int IdAnalisis { get; set; }
           public int IdMetodo { get; set; }
           public string Analisis { get; set; }
           public string Metodo { get; set; }
           public string UnidadMedida { get; set; }
           public string ResultadoDef { get; set; }
           public string LimiteC { get; set; }
           public string LimiteD { get; set; }
           public string Abreviatura { get; set; }
        }
        
        public class CodigoVia
        {
            public int IdProducto { get; set; }
            public string ProductoCodigo { get; set; }
            public string Vias { get; set; }
            public string CodigoInterno { get; set; }
        }

        public class Via
        {
            public int NroViaTemporal;
            public int NroVia;
            public string Presentacion;
            public string PresentacionMuestra;
            public string Muestra;
            public int NumeroMuestra;
            public int IdProducto;
            public string Producto;
            public bool EsAguaPotable;
            public bool EsAguaManantial;
            public string ViaData;
        }
        
        public class ViaResultado
        {
            public string IdProducto { get; set; }
            public int IdAnalisis { get; set; }
            public string UnidMedida { get; set; }
            public string Resultado { get; set; }
            public string CodPrecinto { get; set; }
            public string Muestra { get; set; }
            public string CodigoInterno { get; set; }
            public int NroVia { get; set; }
        }
        
        public class MuestraCls
        {
            public int NroViaTemporal;
            public int NroVia;
            public string Presentacion;
            public string PresentacionMuestra;
            public string Muestra;
            public int NumeroMuestra;
            public int IdProducto;
            public string Producto;
            public bool EsAguaPotable;
            public bool EsAguaManantial;
            public string ViaData;
        }
        
    }
}