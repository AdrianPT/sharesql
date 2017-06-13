using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Globalization;
using System.IO;
using SimpleLogger;
using OPCAutomation;
using Microsoft.SharePoint;
using System.Data.SqlClient;

// License of AutomationOPC Reference https://opcfoundation.org/license/source/1.11/index.html 


namespace PLC_Slm.clases 
{
    public class Make : IJob 
    {

        OPCServer oServer;
        OPCGroup oGroup;
        Array handles;
        public int numItem=0;

        String[] etiqueta;
        String[,] contadorAbsolutoDefiniciones;

        int[,] contadorAbsolutoValores;

        string urlSharepoint;

        string nameSharePointList;
        public int numContadoresAbsolutos = 0;
        public int limMaximoSerie = 0;

        #region Execute
        public void Execute(IJobExecutionContext context)
        {

            string instanciaOPC = readConfig("Geral", "OPC", "");
            string nodoOPC = readConfig("Geral", "Node", "");
            string numeroTags = readConfig("Geral", "Número de tags", "");
            string limiteMaximoSerie = readConfig("Geral", "Limite máximo da serie ", "");
            string numeroContadoresAbsolutos = readConfig("Geral", "Número de contadores absolutos", "");

            urlSharepoint = readConfig("Sharepoint", "Caminho Site", "");
            nameSharePointList = readConfig("Sharepoint", "Nome da Lista", "");

            numItem = Int32.Parse(numeroTags);
            numContadoresAbsolutos = Int32.Parse(numeroContadoresAbsolutos);
            limMaximoSerie = Int32.Parse(limiteMaximoSerie);


            handles = new Array[numItem];
            etiqueta = new String[numItem];

            contadorAbsolutoDefiniciones = new String[numContadoresAbsolutos,3];
            contadorAbsolutoValores = new int[numContadoresAbsolutos, 2];

            for (int i = 1; i <= numItem; i++)
            {
                etiqueta[i - 1] = readConfig("Item"+i, "Nome", "");
            }


            for (int i = 1; i <= numContadoresAbsolutos; i++)
            {

                contadorAbsolutoDefiniciones[i - 1, 0]
                = "ContadorAbsoluto" + i;
                contadorAbsolutoDefiniciones[i-1, 1]
                = readConfig("ContadorAbsoluto" + i, "Tag Valor", "");
                contadorAbsolutoDefiniciones[i - 1, 2]
                = readConfig("ContadorAbsoluto" + i, "Tag Valor Serie", "");
            }


            set_opc(instanciaOPC,nodoOPC);
            leituraTags();
           
	}
        #endregion

        #region OPC Server def
        public void set_opc(String nomeInstancia,String node)
        {
            oServer = new OPCServer();
            oServer.Connect(nomeInstancia, node); // Nodo null puede ser otro 
            oServer.OPCGroups.DefaultGroupIsActive = true;
            oServer.OPCGroups.DefaultGroupDeadband = 0f; 
            oServer.OPCGroups.DefaultGroupUpdateRate = 10; //em ms

            oGroup = oServer.OPCGroups.Add("Grupo 1");
            oGroup.IsSubscribed = false; //suscribir a eventos de mudanças de informação 
            oGroup.OPCItems.DefaultIsActive = false; //el item no necesita estar activo, solo sera actualkizado con el ultimo valor

            //agrega items relativos al grupo siempre es uno mas que los tags
            int[] h = new int[numItem+1];

            //index siempre comienza en 1
            for (int i = 1; i <= numItem; i++)
            {
                h[i] = oGroup.OPCItems.AddItem(etiqueta[i-1], i).ServerHandle; //the handle is a server generated value that we use to reference the item for further operations
            }
            handles = (Array)h;
        }
        #endregion


        #region Tags
        public void leituraTags() //reads device
        {
            System.Array values; //valores
            System.Array errors; //erros
            object qualities = new object(); //qualidade do item
            object timestamps = new object(); //timestamp de leitura opc server

            oGroup.SyncRead((short)OPCAutomation.OPCDataSource.OPCDevice, numItem, ref handles, out values, out errors, out qualities, out timestamps);

            

            Console.WriteLine("Leitura "+ DateTime.Now.ToString("h:mm:ss tt"));
            Console.WriteLine("                                 ");





















            // Contadores Absolutos
            for (int x = 0; x <= numContadoresAbsolutos-1; x++)
            {
                


            //index siempre comienza en 1
            for (int i = 1; i <= numItem; i++)
            {
                
               // String aux = etiqueta[i - 1] + " : {0}";
              //  Console.WriteLine(aux, ((int)values.GetValue(i)));

                // InsertSharepoint(siteURL, nomeLista, impNomeTag, valor);
                if(etiqueta[i - 1].Equals( contadorAbsolutoDefiniciones[x,1])) {

                    
                    contadorAbsolutoValores[x, 1] += ((int)values.GetValue(i));
                        

                       
                }
                else if (etiqueta[i - 1].Equals(contadorAbsolutoDefiniciones[x, 2]))
                {
                    /////////// LIMITE DE 20 ////////////////
                    contadorAbsolutoValores[x, 1] += ((int)values.GetValue(i)) * limMaximoSerie;
                }

           


            } // ^Fin de i


            Console.WriteLine(contadorAbsolutoDefiniciones[x, 0] + " " + contadorAbsolutoValores[x, 1]);



            // 18-01-2017 Checkpoint SQL Sharepoint config.ini
            
             if(readConfig("Mode", "Sharepoint", "").Equals("ON")) InsertSharepoint(urlSharepoint, nameSharePointList, contadorAbsolutoDefiniciones[x, 0], contadorAbsolutoValores[x, 1]);
             if (readConfig("Mode", "Sql", "").Equals("ON")) InsertSQL(contadorAbsolutoDefiniciones[x, 0], contadorAbsolutoValores[x, 1]);


            }            
            Console.WriteLine("");
            Console.WriteLine("");
        }
        #endregion
















        #region ReadConfig
        public string readConfig(string MainSection, string key, string defaultValue)
        {
            string urlConfig = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            urlConfig = urlConfig + "\\config.ini";

            IniFile inif = new IniFile(urlConfig);
            string value = "";

            value = (inif.IniReadValue(MainSection, key, defaultValue));
            return value;
        }
        #endregion











        #region InsertSQL

        protected void InsertSQL(String impNomeTag, int valorReal)
        {
            String server = readConfig("SQL", "DataSource", "");
            String database = readConfig("SQL", "Database", "");
            String UID = readConfig("SQL", "UID", "");
            String PWD = readConfig("SQL", "PWD", "");
            String Tabela = readConfig("SQL", "Tabela", "");
            String ColNomeEtiqueta = readConfig("SQL", "ColNomeEtiqueta", "");
            String ColTimeStamp = readConfig("SQL", "ColTimeStamp", "");
            String ColValorReal = readConfig("SQL", "ColValorReal", "");
            String ColValorVirtual = readConfig("SQL", "ColValorVirtual", "");
            int valorVirtual = 0;
            String connectionSt = "Server=" + server + ";Database=" + database+";User Id=" +  UID +";Password=" + PWD + ";";

            using (SqlConnection connection = new SqlConnection(connectionSt))
            {
               
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = connection;           
                    command.CommandType = CommandType.Text;



                    command.CommandText = "INSERT into " + Tabela + " (" + ColTimeStamp + ", " + ColNomeEtiqueta + ", " + ColValorReal + ", " +  ColValorVirtual + ") VALUES (@TimeStamp, @NomeEtiqueta, @valorReal,@valorVirtual)";


                    valorVirtual = valorVirtualSQL(connectionSt, Tabela, ColValorReal, ColValorVirtual, ColNomeEtiqueta,ColTimeStamp, impNomeTag,valorReal); 
                    
                    
                    command.Parameters.AddWithValue("@TimeStamp", DateTime.Now);
                    command.Parameters.AddWithValue("@NomeEtiqueta", impNomeTag);
                    command.Parameters.AddWithValue("@valorReal", valorReal);
                    command.Parameters.AddWithValue("@valorVirtual", valorVirtual);

                    try
                    {
                        connection.Open();
                        int recordsAffected = command.ExecuteNonQuery();
                    }
                    catch (SqlException ex)
                    {
                        Console.WriteLine(ex);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }
        }

        #endregion





        #region InsertSharepoint

        protected void InsertSharepoint(
       String siteURL, String nomeLista,
       String impNomeTag, int valorReal)
        {
            int valorVirtual = 0;

            try
            {
                ClientContext clientContext = new ClientContext(siteURL);
                SP.List oList = clientContext.Web.Lists.GetByTitle(nomeLista);

                /*
                 *  Recuerda que el valor virtual será:
                 *  ValorVirtual[Actual] = ValorReal [Actual] - ValorReal [UltimoItem] + ValorVirtual [UltimoItem]
                 */


                valorVirtual = vv(siteURL, nomeLista, valorReal, impNomeTag); 


                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);
              
                

                oListItem["Timestamp"] = DateTime.Now;
                oListItem["NomeEtiqueta"] = impNomeTag;
                oListItem["ValorReal"] = unchecked((int)valorReal);
                oListItem["ValorVirtual"] = unchecked((int)valorVirtual);
                

                oListItem.Update();

                clientContext.ExecuteQuery();
            }

            catch (Exception se)
            {
                Console.WriteLine(se);
                SimpleLog.Log(se);
            }


        }
   
        #endregion



       public int vv(String siteURL,String nomeLista,int valorReal,String nomeEtiqueta)
       {
           int valorVirtual =0;
           ClientContext context = new ClientContext(siteURL);
           List list = context.Web.Lists.GetByTitle(nomeLista);
           CamlQuery query = new CamlQuery();
           query.ViewXml = /*"<View><RowLimit>1</RowLimit><OrderBy><FieldRef Name='Timestamp' Ascending='False' /> </OrderBy></View>";*/
               "<View><Query><Where>  <Eq><FieldRef Name='NomeEtiqueta' /><Value Type='Text'>"+nomeEtiqueta+"</Value></Eq>"
             + "</Where>"
             + "<OrderBy><FieldRef Name='Created' Ascending='False' /></OrderBy>"
             + "</Query><RowLimit>1</RowLimit></View>";
           ListItemCollection items = list.GetItems(query);
           context.Load(list);
           context.Load(items);
           context.ExecuteQuery();

           foreach (ListItem item in items)
           {
               
               int aux1 = Convert.ToInt32(item["ValorReal"]);
               int aux2 = Convert.ToInt32(item["ValorVirtual"]);
               valorVirtual = valorReal - aux1 + aux2;
               
               /* ValorVirtual[Actual] = ValorReal [Actual] - ValorReal [UltimoItem] + ValorVirtual [UltimoItem] */
              // Console.WriteLine(valorVirtual);

               }

           return  valorVirtual;
       }




      
       public int valorVirtualSQL(String con, String Tabela,String ColValorReal, String ColValorVirtual,String ColNomeEtiqueta,String ColTimeStamp, String impNomeTag, int valorReal) {



           SqlConnection conn = new SqlConnection(con);
           conn.Open();
                   
           int valorVirtual= 0;

           SqlCommand command =new SqlCommand(
           "SELECT TOP 1 "+ColValorReal+","+ColValorVirtual+" from ["+Tabela+"] where "+ColNomeEtiqueta+"=@nome order by "+ColTimeStamp+" desc", conn);
           
           command.Parameters.AddWithValue("@nome", impNomeTag);
   

        


           // int result = command.ExecuteNonQuery();
           using (SqlDataReader reader = command.ExecuteReader())
           {
               if (reader.Read())
               {
                   // Console.WriteLine("SQL VV: " + String.Format("{0}", reader[ColValorReal]));
                   int aux1 = Convert.ToInt32(reader[ColValorReal]);
                    int aux2 = Convert.ToInt32(reader[ColValorVirtual]);
                    valorVirtual = valorReal - aux1 + aux2;
               }
           }
          
           conn.Close();
       

           return valorVirtual;
       
       }


       




    }


   

}
