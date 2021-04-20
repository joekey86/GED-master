using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json;
using Microsoft.Extensions.Logging;

namespace GED
{
    class SPFunctions
    {

        public static void getStatus (string webUrl, int Id , string listName,ILogger log)
        {
            using (ClientContext ctx = SPConnection.GetSPOLContext(webUrl))
            {
                try
                {
                    string envoi = "";
                    string status = "";
                    string Cycle = "";
                    FieldUserValue[] emptyuser = null;
                    List<String> CibleIDs = new List<string>();


                    List itemList = ctx.Web.Lists.GetByTitle(listName);
                    ListItem item = itemList.GetItemById(Id);
                    ctx.Load(item);
                    ctx.ExecuteQuery();
                    
                    
                    status = item["Etat"].ToString();
                    envoi = item["Envoyer_x0020_le_x0020_document_x0020_pour_x0020_v_x00e9_rification_x002f_validation"].ToString();
                    Cycle = item["Cycle_x0020_de_x0020_vie"].ToString();
                    string OldStatus = item["AncienEtat"] != null ? item["AncienEtat"].ToString() : string.Empty;
                    log.LogInformation("Old Status " + OldStatus);
                    log.LogInformation("Status " + status);
                    if (!string.Equals(status, OldStatus) && item["Cat_x00e9_gories"] !=null)
                    {
                        if(item["R_x00e9_f_x00e9_rence"] == null)
                        {
                            TaxonomyFieldValueCollection tax = item["Cat_x00e9_gories"] as TaxonomyFieldValueCollection;
                            TaxonomyFieldValue typeDoc = item["Type_x0020_de_x0020_document"] as TaxonomyFieldValue;
                            item["R_x00e9_f_x00e9_rence"] = SetReference(ctx, tax, typeDoc,item.Id) ;
                        }
                        log.LogInformation("Enterin Update");

                        item["AncienEtat"] = status;
                       // item.SystemUpdate();
                        ResetBreakRoleInheritance(ctx, item);

                        switch (status)
                        {
                            case "Brouillon":
                                SendEmail(webUrl, (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], "GED - Un document à été créé", 1, (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_R_x00e9_dacteur_x0028_s_x0029_"], item);
                                SPPermissionAuteur(ctx, item, "modify", (FieldUserValue)item["Author"], false);
                                SPPermission(ctx, item, "modify", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);                         
                                UpdateReceivedEmail(item, "Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_R_x00e9_dacteur_x0028_s_x0029_", (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_R_x00e9_dacteur_x0028_s_x0029_"], (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"]);
                               
                                //  UpdateReceivedEmail(ctx, "Redacteur", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], listName, Id);

                                if (envoi == "True" && Cycle.Contains("Rédaction/Validation"))
                                {
                                    item["Etat"] = "En attente de validation";
                                    //  item.Update();

                                }
                                if (envoi == "True" && !Cycle.Contains("Rédaction/Validation"))
                                {
                                    item["Etat"] = "En attente de vérification";
                                    //item.Update();
                                }
                               // item.SystemUpdate();
                               
                                break;
                            case "En attente de vérification":
                                SPPermissionAuteur(ctx, item, "read", (FieldUserValue)item["Author"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "modify", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                                SendEmail(webUrl, (FieldUserValue[])item["V_x00e9_rificateurs"], "GED - Demande de vérification d'un document", 2, (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_V_x00e9_rificateur_x0028_s_x0029_"], item);
                                UpdateReceivedEmail(item, "Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_V_x00e9_rificateur_x0028_s_x0029_", (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_V_x00e9_rificateur_x0028_s_x0029_"], (FieldUserValue[])item["V_x00e9_rificateurs"]);
                                //  UpdateReceivedEmail(ctx, "Verificateur", (FieldUserValue[])item["V_x00e9_rificateurs"], listName, Id);

                                break;
                            case "En attente de validation":
                                SPPermissionAuteur(ctx, item, "read", (FieldUserValue)item["Author"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                                SPPermission(ctx, item, "modify", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                                SendEmail(webUrl, (FieldUserValue[])item["Validateur_x0028_s_x0029_"], "GED - Un document est en attente de validation", 6, (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_aux_x0020_V_x00e9_rificateur_x0028_s_x0029_"], item);
                                UpdateReceivedEmail(item, "Email_x0020_envoy_x00e9__x0020_aux_x0020_V_x00e9_rificateur_x0028_s_x0029_", (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_aux_x0020_V_x00e9_rificateur_x0028_s_x0029_"], (FieldUserValue[])item["Validateur_x0028_s_x0029_"]);
                                // UpdateReceivedEmail(ctx, "Validateur", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], listName, Id);
                                break;
                            case "En attente de publication":
                                SPPermissionAuteur(ctx, item, "read", (FieldUserValue)item["Author"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                                break;
                            case "Publié":
                                SPPermissionAuteur(ctx, item, "read", (FieldUserValue)item["Author"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                                //Cible individuel info
                                SendEmail(webUrl, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], "GED - Nouveau document publié", 3, (FieldUserValue[])item["Email_x0020_Cible_x0020_indiv_x0020_info"], item);
                                UpdateReceivedEmail(item, "Email_x0020_Cible_x0020_indiv_x0020_application", (FieldUserValue[])item["Email_x0020_Cible_x0020_indiv_x0020_application"], (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"]);
                                // UpdateReceivedEmail(ctx, "CibleIndvInfo", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], listName, Id);
                                string emailsToSend = string.Empty;

                                //cible individuel application
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                                SendEmail(webUrl, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], "GED - Nouveau document publié", 4, (FieldUserValue[])item["Email_x0020_Cible_x0020_indiv_x0020_application"], item);
                                UpdateReceivedEmail(item, "Email_x0020_Cible_x0020_indiv_x0020_info", (FieldUserValue[])item["Email_x0020_Cible_x0020_indiv_x0020_info"], (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"]);
                                // UpdateReceivedEmail(ctx, "Email_x0020_Cible_x0020_indiv_x0020_application"", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], listName, Id);
                                ClientContext ctx1 = SPConnection.GetSPOLContext(webUrl);
                                //cible collective info
                                ApplyCibCollectiveProcess(ctx, ctx1, item, "Email_x0020_Cible_x0020_indiv_x0020_info", "Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information", false);
                                //SendEmail(ctx, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], "Subject", "from", "body", emptyuser);

                                //cible collective application
                                ApplyCibCollectiveProcess(ctx, ctx1, item, "Email_x0020_Cible_x0020_indiv_x0020_application", "Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application", true);
                                ctx1.Dispose();
                                //SendEmail(ctx, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], "Subject", 1, emptyuser,item);


                                break;
                            case "En attente de révision":                               
                                SPPermissionAuteur(ctx, item, "modify", (FieldUserValue)item["Author"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                                SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                                SendEmail(webUrl, (FieldUserValue[])item["Author"], "GED - Document en attente de révision", 5, emptyuser, item);
                                if (item["Passer_x0020_en_x0020_publi_x00e9_"].ToString() == "Yes")
                                {
                                    item["Etat"] = "Publié";
                                    // UpdateStatus(ctx, listName, Id);
                                }
                              
                                break;
                              
                                
                        }
                        item.SystemUpdate();
                        ctx.ExecuteQuery();
                    }
                }
                catch(Exception ex)
                {
                    log.LogInformation(ex.ToString());
                    throw ex;
                }
            }
        }

        public static void SPPermission(ClientContext ctx, ListItem item , string role , FieldUserValue[] users, bool createADL)
        {
           // using (ClientContext ctx = SPConnection.GetSPOLContext(ctx1.Url))
            {
               // ListItem item = ctx.Web.Lists.GetByTitle("GED").GetItemById(item1.Id);
                //ctx.Load(item);
                foreach (FieldUserValue user in users)
                {
                    
                    User userpermission = ctx.Web.SiteUsers.GetById(user.LookupId);
                    ctx.Load(userpermission);
                    ctx.ExecuteQuery();
                    // item.BreakRoleInheritance(true, true);
                    RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx);
                    if (createADL)
                    {
                        AddAccusseDeLecture(ctx, item.Id, item["Title"].ToString(), userpermission);
                    }
                    if (role == "modify")
                    {
                        collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type

                    }
                    else if (role == "read")
                    {
                        collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
                    }
                   // if(item.RoleAssignments.GetByPrincipalId(user.LookupId) !=null)
                       // item.RoleAssignments.GetByPrincipalId(user.LookupId).DeleteObject();
                    item.RoleAssignments.Add(userpermission, collRoleDefinitionBinding);
                }
             //   item.SystemUpdate();
              //  ctx.ExecuteQuery();
            }   
        }
        public static void ResetBreakRoleInheritance(ClientContext ctx, ListItem item)
        {
            item.ResetRoleInheritance();
            item.BreakRoleInheritance(false, true);
            item.SystemUpdate();
            ctx.ExecuteQuery();
        }
        public static void SPPermissionByUser(ClientContext ctx, ListItem item, string role, User usr)
        {

          //  User user = ctx.Web.EnsureUser(userEmail);
           // User userpermission = ctx.Web.SiteUsers.GetById(user.LookupId);
            //  item.BreakRoleInheritance(true, true);
            RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx);
           
            if (role == "modify")
            {
                collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type

            }
            else if (role == "read")
            {
                collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
            }
            item.RoleAssignments.Add(usr, collRoleDefinitionBinding);


        }
        public static void SPPermissionAuteur(ClientContext ctx, ListItem item, string role, FieldUserValue user, bool createADL, string taxLabel = null, string termGuid = null)
        {
            
            
                User userpermission = ctx.Web.SiteUsers.GetById(user.LookupId);
            //  item.BreakRoleInheritance(true, true);
            ctx.Load(userpermission);
                RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx);
                if (createADL)
                {
                    AddAccusseDeLecture(ctx, item.Id, item["Title"].ToString(), userpermission,taxLabel,termGuid);
                }
                if (role == "modify")
                {
                    collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type

                }
                else if (role == "read")
                {
                    collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
                }
                item.RoleAssignments.Add(userpermission, collRoleDefinitionBinding);
            

        }
        public static void AddAccusseDeLecture (ClientContext ctx , int docID ,  string docName ,User lecteur,string taxLabel =null,string termGuid = null )
        {

            List itemList = ctx.Web.Lists.GetByTitle("Accusés de lecture");
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem itemtoADD = itemList.AddItem(itemCreateInfo);
            itemtoADD["Title"] = docName;
            itemtoADD["Lecteur"] = lecteur;
            if (!string.IsNullOrEmpty(taxLabel))
            {
                var termValue = new TaxonomyFieldValue();
                termValue.Label = taxLabel;
                termValue.TermGuid = termGuid;
                termValue.WssId = -1;

                itemtoADD["UF_x0020_du_x0020_lecteur"] = termValue;
                    }
            itemtoADD["Document"] = docID;
            itemtoADD["Document_x0020_lu"] = false;
            itemtoADD["Commentaires_x0020__x0028_inform"] = "";


            itemtoADD.SystemUpdate();
            ctx.ExecuteQuery();
        }
        public static void ApplyCibCollectiveProcess(ClientContext ctx, ClientContext ctx1, ListItem item, string emailCibField, string cibleCollectiveField, bool createAL)
        {
            List<String> CibleIDs = new List<string>();
            CibleIDs = GetTaxonomiesId(item, cibleCollectiveField);
            FieldUserValue[] ciblInd = (FieldUserValue[])item[emailCibField];
            List<FieldUserValue> lstCibIndId = new List<FieldUserValue>();
            int cibIndCount = 0;
            if (ciblInd != null)
                cibIndCount = ciblInd.Length;
            foreach (string CibleID in CibleIDs)
            {
                Task<string> azFun = CallEmailAzAsync("wf_Get-Emails-from-CH-Pole-UF", CibleID);


                ListUserDetails lstUserDetails = JsonConvert.DeserializeObject<ListUserDetails>(azFun.Result);
                foreach (UserDetails usd in lstUserDetails.usersDetails)
                {
                    User usr;
                    try
                    {
                        usr = ctx1.Web.EnsureUser(usd.UserName);
                        ctx1.Load(usr);
                        ctx1.ExecuteQuery();
                        FieldUserValue usrValue = new FieldUserValue();
                        usrValue.LookupId = usr.Id;
                        if (ciblInd.Where(a=> a.Email == usrValue.Email) == null)
                        {

                            lstCibIndId.Add(usrValue);
                            SPPermissionAuteur(ctx, item, "read", usrValue, createAL, usd.UFDefaultLabel, usd.UFMetadataID);

                        }
                        else
                        {
                            SPPermissionAuteur(ctx, item, "read", usrValue, false);
                        }
                    }
                    catch { }

                }
            }
            if (lstCibIndId.Count() > 0)
            {
                int counter = 0;
                int arrLength = lstCibIndId.Count() + ciblInd.Length;
                FieldUserValue[] newCibInd = new FieldUserValue[arrLength];
                foreach (FieldUserValue usV in ciblInd)
                {
                    newCibInd[counter] = usV;
                    counter++;
                }
                foreach (FieldUserValue usV in lstCibIndId)
                {
                    newCibInd[counter] = usV;
                    counter++;
                }
                item[emailCibField] = newCibInd;
                
            }
        }
        public static void SendEmail(string webUrl, FieldUserValue[] users , string Subject, int index , FieldUserValue[] userReceived,ListItem item)
        {

            using (ClientContext ctx = SPConnection.GetSPOLContext(webUrl))
            {
                foreach (FieldUserValue user in users)
            {
                bool notcontain = true; 
                if(userReceived != null)
                {
                    foreach (FieldUserValue recuser in userReceived)
                    {
                        if (user.Email == recuser.Email)
                        {
                            notcontain = false;
                        }
                    }
                }

                    if (notcontain)
                    {

                        List<string> usersEmail = new List<string> { };
                        usersEmail.Add(user.Email.ToString());
                        string body = "";
                        if (index == 4)
                        {
                            body = EmailBody(index, item, user.Email.ToString(), users);
                        }
                        else
                        {
                            body = EmailBody(index, item, user.Email.ToString());
                        }
                        try
                        {
                            ///jke risk of dispose
                            //  using (ctx)
                            {

                                var emailProperties = new EmailProperties();
                                //Email of authenticated external user
                                emailProperties.To = usersEmail;
                                emailProperties.From = "process@ghtpdfr.fr";
                                emailProperties.Body = body;
                                emailProperties.Subject = Subject;
                                //emailProperties.CC = cc;
                                Utility.SendEmail(ctx, emailProperties);



                                ctx.ExecuteQuery();



                            }
                        }
                        catch (Exception ex)
                        {



                        }

                    }
                }
            }
         }
        public static void UpdateStatus(ClientContext ctx , string listName , int Id)
        {
            List itemList = ctx.Web.Lists.GetByTitle(listName);
            ListItem item = itemList.GetItemById(Id);       
            item["Etat"] = "Publié";
            item.Update();
            ctx.ExecuteQuery();
        }

        public static void UpdateReceivedEmail(ListItem item,string fieldName, FieldUserValue[] usersToUpdate, FieldUserValue[] usersToAdd )
        {
            if (usersToAdd != null)
            {
                List<FieldUserValue> users = new List<FieldUserValue>();
                foreach (FieldUserValue usr in usersToAdd)
                {
                    if (usersToUpdate== null || usersToUpdate.Length == 0 || usersToUpdate.Where(a => a.Email == usr.Email) == null)
                    {
                        users.Add(usr);
                    }
                }
                int userToUpdateLength = usersToUpdate != null ? usersToUpdate.Length : 0;
                int totalCount = userToUpdateLength + users.Count;
                if (totalCount > 0)
                {
                    FieldUserValue[] newUsers = new FieldUserValue[totalCount];
                    if (usersToUpdate != null)
                        usersToUpdate.CopyTo(newUsers, 0);
                    else
                    if(users.Count >0)
                    {
                        users.CopyTo(newUsers, 0);
                    }
                    if (users.Count > 0 && usersToUpdate !=null && usersToUpdate.Length > 0)
                        users.CopyTo(newUsers, usersToUpdate.Length);
                    

                    item[fieldName] = newUsers;
                }
              }
           // List itemList = ctx.Web.Lists.GetByTitle(listName);
           // ListItem item = itemList.GetItemById(Id);
            ////if (list == "Redacteur")
            ////{

            ////    item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_R_x00e9_dacteur_x0028_s_x0029_"] = users;


            ////}
            ////else if (list == "Verificateur")
            ////{
            ////    item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_V_x00e9_rificateur_x0028_s_x0029_"] = users;

            ////}
            ////else if (list == "Validateur")
            ////{
            ////    item["Email_x0020_envoy_x00e9__x0020_aux_x0020_V_x00e9_rificateur_x0028_s_x0029_"] = users;

            ////}
            ////else if (list == "CibleIndvApp")
            ////{

            ////    item["Email_x0020_Cible_x0020_indiv_x0020_application"] = users;

            ////}
            ////else if (list == "CibleIndvInfo")
            ////{
            ////    item["Email_x0020_Cible_x0020_indiv_x0020_info"] = users;

            ////}
            ////item.Update();
            ////ctx.ExecuteQuery();
        }

        public static string GetAppSetting(string key)
        {
            return Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process);
        }


        public static async Task<string> CallEmailAzAsync(string wfName, string termGuid)
        {
            string result = string.Empty;
            string body = "{\"managedmetadataID\": \"" + termGuid + "\"}";
            string url = "https://prod-08.francecentral.logic.azure.com:443/workflows/9085e566a0c64c6ca3a48811f215d975/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IOg706o5XUEfo3RIjLQLOaSZi3uqoCKxjqPFc8y6Vwk";
            //string url = GetAppSetting("wf_Get-Emails-from-CH-Pole-UF");
            using (var httpClient = new HttpClient())
            {
                var content = new StringContent(body, Encoding.UTF8, "application/json");
                var response =  httpClient.PostAsync(url, content).Result;
                result = response.Content.ReadAsStringAsync().Result;
            }
            return result;
        }
        //public static string GetTaxonomyId(ListItem item, string fieldName)
        //{

        //    TaxonomyFieldValue taxFieldValue = item[fieldName] as TaxonomyFieldValue;
        //    return taxFieldValue.TermGuid;
        //}
        public static List<string> GetTaxonomiesId(ListItem item, string fieldName)
        {
            List<string> Ids = new List<string>();
           // string er = item[fieldName].ToString();
            try
            {
                System.Collections.Generic.Dictionary<System.String, System.Object> sds = item[fieldName] as System.Collections.Generic.Dictionary<System.String, System.Object>;
                TaxonomyFieldValueCollection taxFieldValues = sds.ElementAt(1).Value as TaxonomyFieldValueCollection;
                object[] DSD = sds.ElementAt(1).Value as object[];
                foreach (Dictionary<System.String, System.Object> dic in DSD)
                {

                    Ids.Add(dic["TermGuid"].ToString());

                }
            }
            catch(Exception ex)
            {
                TaxonomyFieldValueCollection taxFieldValues = item[fieldName] as TaxonomyFieldValueCollection;

                foreach (TaxonomyFieldValue taxFieldValue in taxFieldValues)
                {

                    Ids.Add(taxFieldValue.TermGuid);

                }
            }
            return Ids;
        }

        public static string EmailBody(int index, ListItem item, string useremail,FieldUserValue[] users = null)
        {
            string body = "";
            string fileUrl = "https://ghtpdfr.sharepoint.com/" + item["FileRef"].ToString();
            if (index == 1)
            {
                body = @"Bonjour,
                        Un document à été créé dans la bibliothèque GED: <br/><br/>
                            " +
                         "<a href='"+ fileUrl + "' >"+ fileUrl + "</a>";

            }
            else if (index == 2)
            {
                body = @"Bonjour,
                        Un document demande à être vérifié dans la bibliothèque GED: <br/><br/><br/>
                        <a href ='"+ fileUrl + "' > "+ fileUrl + @"</a>; <br/><br/>

                        Si vous jugez que le document est vérifié et doit passer en validation, veuillez cliquer sur ce <a href ='https://prod-30.francecentral.logic.azure.com/workflows/6ba778559279416580dd5c3cfdef3213/triggers/manual/paths/invoke/" + useremail + "/" + item.Id + "/" + item["Etat"].ToString() + "/True?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BiOqRLe2hB-pDrBG-hWVk2KMdiD_4wuEE96hiZVEWws' >lien </a>" +
                        @" <br/><br/>
                        
                        Sinon, veuillez cliquer sur ce <a href ='https://prod-30.francecentral.logic.azure.com/workflows/6ba778559279416580dd5c3cfdef3213/triggers/manual/paths/invoke/" + useremail + "/" + item.Id + "/" + item["Etat"].ToString() + "/false?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BiOqRLe2hB-pDrBG-hWVk2KMdiD_4wuEE96hiZVEWws' > lien </a>";


            }
            else if (index == 3)
            {

                body = @"Bonjour,
                            Un document a été publié dans la bibliothèque GED:<br/><br/>                        
                              <a href ='" + fileUrl + "' > " + fileUrl + "</a>";



            }
            else if(index == 4)
            {
                body = @"Bonjour,
                        Un document a été publié dans la bibliothèque GED:<br/><br/>
                        <a href ='" + fileUrl + "' > " + fileUrl + @"</a><br/>


                        Veuillez cliquer sur ce <a href ='https://prod-10.francecentral.logic.azure.com/workflows/39c7e411b3bb4c7f9bd122bbffe5f170/triggers/manual/paths/invoke/" + users.First().LookupId + "/" + item.Id + "?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=mG931Fl_dDU6s02Z5XJ7wdY-DfmPrj-t6-chxiLcs6A' >lien</a> pour valider la lecture du document";

                        

            }
            else if(index == 5)
            {
                body = @"Bonjour,
                        Votre document est en attente de révision:<br/><br/>
                        <a href ='" + fileUrl + "' > " + fileUrl + @"</a>";
                                           
            }
            else if (index == 6)
            {
                body = @"Bonjour,
                        Un document demande à être validé dans la bibliothèque GED <br/><br/><br/>
                        <a href ='" + fileUrl + "' > " + fileUrl + @"</a>; <br/><br/>

                        Si vous voulez valider le document, veuillez cliquer sur ce <a href ='https://prod-30.francecentral.logic.azure.com/workflows/6ba778559279416580dd5c3cfdef3213/triggers/manual/paths/invoke/" + useremail + "/" + item.Id + "/" + item["Etat"].ToString() + "/True?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BiOqRLe2hB-pDrBG-hWVk2KMdiD_4wuEE96hiZVEWws' >lien </a>" +
                        @" <br/><br/>
                        
                        Sinon, veuillez cliquer sur ce <a href ='https://prod-30.francecentral.logic.azure.com/workflows/6ba778559279416580dd5c3cfdef3213/triggers/manual/paths/invoke/" + useremail + "/" + item.Id + "/" + item["Etat"].ToString() + "/false?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BiOqRLe2hB-pDrBG-hWVk2KMdiD_4wuEE96hiZVEWws' > lien </a>";


            }
            return body;
        }
        public static string SetReference(ClientContext ctx, TaxonomyFieldValueCollection tax, TaxonomyFieldValue typeDoc, int Id)
        {

  
                                          
TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            taxonomySession.UpdateCache();
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            ctx.Load(termStore,
                termStoreArg => termStoreArg.WorkingLanguage,
                termStoreArg => termStoreArg.Id,
                termStoreArg => termStoreArg.Groups.Include(
                    groupArg => groupArg.Id,
                    groupArg => groupArg.Name
                )
            );            
            Guid catTermId = new Guid(tax.ElementAt(0).TermGuid);
            TermSet catTermSet = termStore.GetTermSet(new Guid("6792f6c1-20ec-4e10-a1e8-a2c04f2906ec"));
            Term catTerm = catTermSet.GetTerm(catTermId);
            ctx.Load(catTermSet);
            ctx.Load(catTerm);
            ctx.Load(catTerm.Labels);
           // ctx.Load(catTerm.Parent.Labels);
            Guid docTermId = new Guid(typeDoc.TermGuid);
            TermSet docTermSet = termStore.GetTermSet(new Guid("40ae95fa-353f-4154-a574-65f7297286ca"));
            Term docTerm = docTermSet.GetTerm(docTermId);
            ctx.Load(docTermSet);          
            ctx.Load(docTerm);       
            ctx.Load(docTerm.Labels);
            ctx.ExecuteQuery();
            string reference = catTerm.Labels[1].Value + "/" + docTerm.Labels[1].Value + "/" + Id;
            //if(catTerm.Parent.Labels.Count > 1)
            //{
            //   reference = catTerm.Parent.Labels[1].Value + "/" + catTerm.Labels[1].Value + "/" + docTerm.Labels[1].Value + "/" + Id;
            //}
            return reference;
        }
    }

}
