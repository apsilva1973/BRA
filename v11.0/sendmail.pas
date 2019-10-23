  EmailDotNet : TEmailNet;

                  EmailDotNet := TEmailNet.Create(ExtractFilePath(Application.ExeName));
                  try


                       result := EmailDotNet.Enviar_Email(mailInfo.smtpObj.HostName,
                                                          mailinfo.smtpObj.PortNumber,
                                                          mailInfo.smtpObj.LoginName,
                                                          mailInfo.smtpObj.PassWord,
                                                          mailInfo.smtpObj.AddressName,
                                                          emailLst,
                                                          emailCC,
                                                          emailCCO,
                                                          'N',
                                                          myFiles,
                                                          qryMsgs.FieldByName('Assunto').AsString + ' - ' +
                                                          qryMsgs.FieldByName('COD_DEP_Remetente').AsString +
                                                          ' ('+ qryMsgs.FieldByName('Num_Sisbacen').AsString  + ' ' +	')',
                                                          myheader.text,
                                                          'N',
                                                          true,
                                                          'S',
                                                          vim830Global.totalCaixasPoremail,
                                                          Erro);

