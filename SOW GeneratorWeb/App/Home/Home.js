// <reference path="../App.js" />

(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        var $loading = $('#status');
        var $spinner = $('#spinner');
        var inc = 0;
        $(document)
        .ajaxStart(function () {
            // When Ajax calls start, show spinner.
            $("#statcomp").html("");
            $spinner.show();
            progress(0, $loading);
            $loading.show();
            progress(15, $loading);
            inc = 20;
        })
        .ajaxComplete(function () {
            progress(inc, $('#status'));
            inc = inc + 5;
        })
        .ajaxStop(function () {
            // When all AJAX calls are complete, stop spinner
            setTimeout(function () {
                //$loading.fadeOut("slow");
                progress(100, $loading);
                $spinner.fadeOut("slow");
                $loading.fadeOut("slow");
                $("#statcomp").html("Generation Complete.");
            }, 8000);
        })
        .ready(function () {
            app.initialize();
            
            // set variable for type of SOW generation (i.e. CERQ or PID)
            var gentype;

            // On load, show CERQ search
            if ($('#btncerq').is(':checked')) {
                showcerq();
            };
            
            // Show/hide proper inputs based on radio button selection by user
            $('input[type = "radio"]').click(function(){
                if ($('#btncerq').is(':checked')) {
                    showcerq();
                }

                if ($('#btnnocerq').is(':checked')) {
                    shownocerq();
                }
            });

            // Show CERQ search, hide PID search
            function showcerq() {
                $('#bycerq').show();
                $('#byPID').hide();
                gentype = "CERQ";
            }

            // Show PID Search, hide CERQ search
            function shownocerq() {
                $('#byPID').show();
                $('#bycerq').hide();
                gentype = "PID";
            }

            // Initially hide loading spinner
            $loading.hide();
            // Accordion for Settings on task pane
            $("#accordion").accordion({ collapsible: true, active: false });
            
            // Submit button func
            $('#s').submit(function (e) {
                // Clear any previous errors
                document.getElementById('message').innerText = "";

                if (gentype === "CERQ") {
                var cerq = 'CERQ' + $("#cerq").val();
                var prod = $("#prod").val();
                // Begins the process of gathering data and binding to Word
                begin(cerq);
                }

                if (gentype === "PID") {
                    var pid = $("#PID").val();
                    var jcodes = $("#jcodes").val();
                    if (pid.length != 5) {
                        $('#notification-message').css({ "background-color": "#FFBABA", "color": "#D8000C" });
                        app.showNotification('INCORRECT PID!', "Please input a 5-digit PID number.");

                    } else {
                        pid = '000' + pid;
                        beginPID(pid, jcodes);
                    }
                }
                // Prevents default action of the submit button
                e.preventDefault();
            });
        })
    };

    // Global variables
    var model = {
        primsolId: '', primsolname: '', oppId: '', layout: '', vb: '', layoutconfig: '', xmldata: '', salesrep: '', customer: '', docname: '', custcontact: '', contactphone: '', custadd0: '', custadd1: '',
        cerq: '', cewarranty: '', manuwarr: '', pmtemplate: '', pmsecpos: '', ordsec: [], secdata: [], xmlengUT: '', engtraintempl: '', warrpostsatempl: '', ceriumadd0: '', ceriumadd1: '', fullname: '',
        phone: '', locations: [], locacctIds: '', modurl: '', locdata: [], salescodeId: '', userId: '', secondarySC: [], scId: '', secscurl: '', scverb: '', loading: '', noSolOverview: '', pmweeks: '', pid: '',
        acctId: '', oppName: '', oppOwnerId: '', oppOwnerName: '', contactId: '', jcodes: []
    };

    var restURL = {
        crmcustom: 'http://'   
       
    }

    // Generation by CERQ
    var begin = function (cerq) {
        // Use the global variable to store the cerq.
        model.cerq = cerq;
        $("#getacct").html("In progress...");

        // Get quote data from QuoteWerks. When finished, run parseacctdata func
        var getacct = $.getJSON(restURL.crmcustom + 'GetQWDocumentHeader/' + model.cerq + '?callback=?')
            .then(parseacctdata);
    }

    // Parse the response from QuoteWerks.
    var parseacctdata = function (data) {
        $.each(data.results, function (i, d) {
            if (d.DocNo) {
                if (d.SoldToCMOppRecID) {
                    if (d.SoldToCMOppRecID.slice(0, 1) === "{") {
                    model.oppId = d.SoldToCMOppRecID.slice(1, 37);
                    } else {
                        model.oppId = d.SoldToCMOppRecID
                    }
                    // Call func to get Opportunity data from CRM.
                    getOppdata(model.oppId); 
                    model.salesrep = d.SalesRep;
                    model.customer = d.SoldToCompany;
                    model.docname = d.DocName;
                    model.custcontact = d.SoldToContact;
                    model.contactphone = d.SoldToPhone;
                    model.custadd0 = d.SoldToAddress1;
                    model.custadd1 = d.SoldToCity + ", " + d.SoldToState + " " + d.SoldToPostalCode;
                } else {
                    write(" No Opportunity ID in QW. ");
                }
            }
        });
        $("#getacct").html("Finished");
        return model.oppId;
    }

    // Generation by PID
    var beginPID = function (pid, jcodes) {
        // Use the global variable to store the cerq.
        model.pid = pid;
        model.jcodes = jcodes;
        $("#getacct").html("In progress...");

        // Get quote data from QuoteWerks. When finished, run parseacctdata func
        var getopp = $.getJSON(restURL.crmcustom + 'GetOpportunitybyPID/' + model.pid + '?callback=?')
            .then(parseOppData);
    }

    function popdata() {
        // Get and populate JCodes (based on how the user searches), fill in Account info, fill contact info        
        getUserdata(model.userId);
        if (model.cerq) {
            setTimeout(function () { getJCodes(model.cerq) }, 1000);
        } else if (model.pid) {
            getJCodesbyPID();
   }
        setTimeout(function () { buildTOC(); }, 1000);
   }

    function fillcontactinfo() {

        // Set Cerium AE name
        // Need setTimeout or they won't load properly.
        setTimeout(function () { binddata('CustomerContact', 'customercontact', model.custcontact, "text"); }, 100);
        setTimeout(function () { binddata('CustomerContactPhone', 'contactphone', model.contactphone, "text"); }, 100);
        setTimeout(function () { binddata('CustomerAddress0', 'companyaddress0', model.custadd0, "text"); }, 100);
        setTimeout(function () { binddata('CustomerAddress1', 'companyaddress1', model.custadd1, "text"); }, 100);
        setTimeout(function () { binddata('cerq', 'cerq', model.cerq, "text"); }, 100);

        // Get and set date in overview section
        var date = new Date();
        var today = (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
        binddata('TodayDate', 'todaydate', today, "text");
    }

    function fillSCdata() {
        setTimeout(function () { binddata('SolCerWarr', 'solcerwarr', model.cewarranty, "html"); }, 500);
        setTimeout(function () { binddata('SolManuWarr', 'solmanuwarr', model.manuwarr, "html"); }, 500);
        
        // Populate these fields if Solution (Software) Overview section exists. When it doesn't exist don't populate.
        if (model.noSolOverview) {
            setTimeout(function () { binddata('SalesCodeVerb', 'salescodeverb', model.scverb, "html"); }, 500);
            setTimeout(function () { binddata('PrimSolType', 'primsoltype', model.primsolname, "html"); }, 500);
        }
    }

    function bindAcct(acctname) {
        // Check for "HQ" in account name and remove it if it exists.
        $("#fillacct").html("In Progress...");
        var acctlength = acctname.length;
        var start = acctlength - 5;
        var end = acctlength + 1;
        var newacctname = acctname.slice(start, end);
        if (newacctname === " - HQ") {
            acctname = acctname.slice(0, start);
        }
        
        // Updates customXML in document with Customer Name and project name. In turn, the plain text box for Customer is updated in the doc header.
        Office.context.document.customXmlParts.getByNamespaceAsync("AccountHolder", function (result) {
            if (result.status == "failed") {
                write('Custom XML Error: ' + result.error.message);
            } else {
                var xmlType = result.value[0];
                // Update account name in header
                xmlType.getNodesAsync('/ns0:AccountHolder[1]/ns1:Account[1]', function (nodeResults) {
                    nodeResults.value[0].setXmlAsync("<Account xmlns='Account'>" + acctname + "</Account>");
                });
                // UPdate project name in header
                xmlType.getNodesAsync('/ns0:AccountHolder[1]/ns2:Project[1]', function (nodeResults) {
                    nodeResults.value[0].setXmlAsync("<Project xmlns='Project'>" + model.docname + "</Project>");
                });
            }
        });
        $("#fillacct").html("Finished");
}
                             
    // Get Opportunity data from CRM.
    var getOppdata = function (oppId) {
        $("#getOppdata").html("In Progress...");
        var oppUrl = restURL.crmcustom + "GetOpportunity/" + oppId + '?callback=?';
        $.getJSON(oppUrl)
            .done(function (data) {
                if (data.results[0].SalesCodeId) {
                    parseOppData(data);
                } else {
                    write("ERROR! Please add a Sales Code to your Opportunity and rerun the SOW Generator.");
                    $(".loader").fadeOut("slow");
                }
            })
            .fail(
                function () {
                    write(" Opportunity call failed. getOppdata. ");
                }
            )
        $("#getOppdata").html("Finished");
    }

    function parseOppData(data) {
                    $.each(data.results, function (i, d) {
                        model.userId = d.OwningUser_Id;
                        model.salescodeId = d.SalesCodeId;
            // If generated by PID
            if (model.pid) {
                model.acctId = d.AccountId;
                model.customer = d.ParentAccountIdName;
                model.oppId = d.OpportunityId;
                model.docname = d.Name;
                // Get Sold To info from CRM, since we don't have info from QW
                getSoldTo();
            }
                        // Get Sales Code data from Sales Code table in CRM.
                        getsalescodedata();
                        // Build TOC
                        getlocalxml('TOC', 'toc', '../../XMLSnippets/Sections/0_TOC.xml');
                        // Get Secondary Sales Code data in CRM
                        setTimeout(function () { getsecondarySalesCode(); }, 500);
                    });

        return model.oppId;
                }

   var getSoldTo = function() {
        // Get Account data
        var accturl = restURL.crmcustom + "GetAccount/" + model.acctId + "?callback=?";
        $.getJSON(accturl)
            .then(function (data) {
                $.each(data.results, function (i, d) {
                    model.contactId = d.PrimaryContactId;
                    model.custadd0 = d.Address1_Line1;
                    model.custadd1 = d.Address1_City + ", " + d.Address1_StateOrProvince + " " + d.Address1_PostalCode;
                })
            }).then(function () {
                // Get the primary contact info
                getContact();
            })
             .fail(function (jqxhr, textStatus, error) {
                 var err = textStatus + ", " + error;
                 write(" Error message for getSoldTo: " + err);
             });
                }

   var getContact = function(){
       var contacturl = restURL.crmcustom + "GetCrmContact/" + model.contactId + "?callback=?";
        $.getJSON(contacturl)
            .then(function (data) {
                $.each(data.results, function (i, d) {
                    model.custcontact = d._FullName;
                    model.contactphone = d._Telephone1;
                    //model.custadd0 = d._Address1_Line1;
                    //model.custadd1 = d._Address2_City + ", " + d._Address1_StateOrProvince + " " + d.Address1_PostalCode;
                });
            });
        $("#getacct").html("Finished");
    } 
                
    // Get Opportunity owner data
    function getUserdata(userId) {
        $("#getuserdata").html("In Progress...");
        //var userURL = restURL.crm + "SystemUser?$select=FullName,Id,Address1_Telephone1,SiteId_Id&$filter=Id eq '" + userId + "'&$format=json&$callback=?&@authtoken=" + restURL.token;
        var userURL = restURL.crmcustom + "GetSystemUser/" + userId + '?callback=?';
        $.getJSON(userURL)
            .done(function (data) {
                $.each(data.results, function (i, d) {
                    var siteId = d.SiteId;
                    model.fullname = d.FullName;
                    model.phone = d.Address1_Telephone1;
                    getsitedata(siteId);
                })
            setTimeout(function () { binddata('CeriumOwner', 'ceriumowner', model.fullname, "text");}, 1500);
            setTimeout(function () { binddata('CeriumOwnerPhone', 'ceriumownerphone', model.phone, "text"); }, 1500);
            $("#getuserdata").html("Finished");
            })
          .fail(function (jqxhr, textStatus, error) {
              var err = textStatus + ", " + error;
              write(" getlayoutdata: URL failure error: " + err);
          });
    } 

    // Get Site data from CRM.
    function getsitedata(siteId) {
        $("#getsitedata").html("In Progress...");
        var url = restURL.crmcustom + "GetSite/" + siteId + '?callback=?';
        $.getJSON(url)
            .done(function (data) {
                $.each(data.results, function (i, d) {
                    if (d.Address1_Line2) {
                        model.ceriumadd0 = d.Address1_Line1 + " " + d.Address1_Line2;
                        binddata('CerAddress0', 'ceriumaddress0', model.ceriumadd0, "text");
                    } else {
                        model.ceriumadd0 = d.Address1_Line1;
                        //binddata('CerAddress0', 'ceriumaddress0', model.ceriumadd0, "text");
                    }
                    var sitename = d.Name;
                    model.ceriumadd1 = d.Address1_City + ", " + d.Address1_StateOrProvince + " " + d.Address1_PostalCode;
                    //binddata('CerAddress1', 'ceriumaddress1', model.ceriumadd1, "text");
                }); 
            setTimeout(function () { binddata('CerAddress0', 'ceriumaddress0', model.ceriumadd0, "text"); }, 1500);
            setTimeout(function () { binddata('CerAddress1', 'ceriumaddress1', model.ceriumadd1, "text"); }, 1500);

            // Call func to fill in Account Name info
            bindAcct(model.customer);
            $("#getsitedata").html("Finished");
        }); 
    } 

    function getLocations(oppId) {
        $("#getLocations").html("In Progress...");
        var url = restURL.crmcustom + "GetOppLocation/" + oppId + '?callback=?';
        $.getJSON(url)
            .done(function (data) {
                $.each(data.results, function (i, d) {
                    model.locations.push(d.accountid);
                });

                if (model.locations) {
                    var build = new Buildaccturl;
                    setTimeout(function () { build.buildurl(); }, 1000); // zmallahan 7/10/15: reduced to 1000 from 3000
                } else {
                    $("#getLocations").html("Finished");
                }
            });
    }

    // Builds URL to retrieve locaton data
    function Buildaccturl() {
        var length = model.locations.length;
        this.buildurl = function () {
            if (length > 0) {
                for (var i = 0; i < length; i++) {
                    if (i === (length - 1)) {
                        model.locacctIds += model.locations[i];
                    } else {
                        model.locacctIds += model.locations[i] + ",";
                    }
                }
                model.modurl = restURL.crmcustom + "GetAccount/" + model.locacctIds + '?callback=?';
                this.getacctdata();
            }
        };

        // Builds the location data into html and loads into a variable for later binding on the SOW.
        this.getacctdata = function () {
            var html = '';
            $.getJSON(model.modurl).done(function (data) {
                $.each(data.results, function (i, d) {
                    html += '<div>' + d.Name + ' | ' + d.Address1_Line1 + ' | ' + d.Address1_City + ', ' + d.Address1_StateOrProvince + ' | ' + d.Telephone1 + '</div></div>';
                });
               // write("calling binddata for locations");
                binddata('Locations', 'locations', html, "html");
                $("#getLocations").html("Finished");
            })
            .fail(function (jqxhr, textStatus, error) {
                var err = textStatus + ", " + error;
                write(" Error binding Locations onto the SOW. Contact Development Team. ");
            });
        }
    }
    
    // Get JCodes from QuotesWerks when searching by CERQ
    function getJCodes(cerq) {
        $("#getjcodes").html("In Progress...");
        var url = restURL.crmcustom + 'GetJCodes/' + cerq + '?callback=?'
        var promise = $.ajax({
            url: url,
            contentType: "application/json;charset-uf8",
            dataType: "jsonp",
            error: function (xhr, ajaxOptions, thrownError) {
                write(xhr.status);
                write(thrownError);
                write(" An error occured while calling getJCodes. ");
            }
        });
        promise.then(parseJCodes);

    } //end getJCodes

    // Get JCodes from QuoteWerks when searching by PID
    function getJCodesbyPID() {
        $("#getjcodes").html("In Progress...");
        // Replace spaces with commas
        var tempjcodes = [];
        tempjcodes = model.jcodes.replace(/[\n\r]/g, ',');
        // Separate array into smaller chunks. If too long, web service will not return any results.
        // First, split string of JCodes into separate array indices        
        var arrayjcodes = tempjcodes.split(',');
        var jcodes = [];
        var size = 20; // Over 34 JCodes is too many. Let's take 20 at a time
        // Crete the number of buckets necessary to split up the JCodes. Each bucket holds no more than 20.
        var bins = Math.ceil((arrayjcodes.length) / size);
        var i = 0;
        var resp = [];
        // Break out JCodes into groups of 20
        while (i++ < bins) {
            jcodes.push(arrayjcodes.splice(0, size));
        }
        // Set iteration variable
        var iterate = (jcodes.length - 1)

        // Get JCode data from QWs, combine it, then send it to be parsed and bound to the form.
        var def = $.Deferred();
        getData(iterate).done(parseJCodes);
        function getData(n) {
            // Setup deferred object to use promises to wait for data to be returned
            var url = restURL.crmcustom + 'GetMultiJCodes/' + jcodes[n] + "?callback=?";
            // Base Case
            if (n === 0) {
                $.getJSON(url).done(function (data) {
                    resp.push(data);
                    def.resolve(resp);
                });

                return def.promise();
            } else {
                // Recursive case
                $.getJSON(url).done(function (data) {
                    resp.push(data);
                });
            }
            var d = getData(n - 1);
            d.done(function () {
                def.resolve(resp);
            })
            return def.promise();
        };
    }

        // Parse jCode data and build HTML based on specific ItemTypes (QuoteWerks field). 
        // Item types have specific placeholders on the SOW document
        function parseJCodes(data) {
            var eng_html = '', pm_html = '', custres_html = '', eng_UT_html = '', eng_SysAdTr_html = '', warr_postsa_html = '', suppctr_html = '', html = '', eng_oth_html = '', wordlength = '';
            $.each(data, function (i, data) {
                $.each(data.results, function (i, d) {
                    html = '<div><b>' + d.key + '</b></div><div>' + d.value + '</div></br>'
                    if (d.itemtype === "Services_Engineering") {
                        // call funciton to remove CRLF characters that come across in data feed.
                        eng_html += replaceCRLF(d.key, d.value);
                    } else if (d.itemtype === "Services_Engineering_Other") {
                        // call funciton to remove CRLF characters that come across in data feed.
                        eng_oth_html += replaceCRLF(d.key, d.value)
                    } else if (d.itemtype === "ProjectManagement") {
                        if (d.CustomText13) {
                            var template = d.CustomText13.split(".", 1);
                            getPMxml(template);
                        }
                        // If JCodes match criteria for Project Management week duration and
                        // they are greater than 4 (weeks) set variable to be used later to SOW verbiage.
                    }
                    else if (d.manufacturerpartnumber === 'J91055' || d.manufacturerpartnumber === 'J91065') {
                        if (d.QtyTotal > 4) {
                            // Call toWord function from toword.js
                            model.pmweeks = toWords(d.QtyTotal);
                            wordlength = model.pmweeks.length;
                            model.pmweeks = model.pmweeks.slice(0, wordlength - 1);
                        } else {
                            model.pmweeks = "four";
                        }
                    }
                    else if (d.itemtype === "CustomerResponsibilities") {
                        custres_html += html;
                    } else if (d.itemtype === "Services_Engineering_UT") {
                        if (!model.engtraintempl) {
                            getlocalxml('Eng_Training', 'eng_training', '../../XMLSnippets/Engineering Services/BuiltInVB_Eng_UT.xml');
                            model.engtraintempl = 1;
                        }
                        //eng_UT_html += html;
                        eng_UT_html += replaceCRLF(d.key, d.value);
                    } else if (d.itemtype === "Services_Engineering_SysAdTr") {
                        if (!model.engtraintempl) {
                            getlocalxml('Eng_Training', 'eng_training', '../../XMLSnippets/Engineering Services/BuiltInVB_Eng_UT.xml');
                            model.engtraintempl = 1;
                        }
                        eng_SysAdTr_html += replaceCRLF(d.key, d.value);
                    } else if (d.itemtype === "Warranty_Post-SA") {
                        //if (!model.warrpostsatempl) {
                        //    getlocalxml('WarrantyPostSA','warrantypostsa','../../XMLSnippets/Warranty/BuiltInVB_Warr_Post_SA.xml');
                        //    model.warrpostsatempl = 1;
                        //}
                        warr_postsa_html += html;
                    } else if (d.itemtype === "SupportCenter") {
                        suppctr_html += html;
                    }
                });
            });

            // Bind the above html to it's respective placeholder on the SOW document.
            if (eng_html) {
                binddata('EngServices', 'engservices', eng_html, "html");
            }
            if (eng_oth_html) {
                binddata('EngServices_Other', 'engservices_other', eng_oth_html, "html");
            }
            if (custres_html) {
                setTimeout(function () { binddata('CustRespJcode', 'custrespjcode', custres_html, "html"); }, 500);
            }
            if (eng_UT_html) {
                //needs to wait for above (eng_html) to bind data
                setTimeout(function () { binddata('Services_Engineering_UT', 'Services_Engineering_UT', eng_UT_html, "html"); }, 2500);
            }
            if (eng_SysAdTr_html) {
                setTimeout(function () { binddata('Services_Engineering_SysAdTr', 'services_engineering_sysadtr', eng_SysAdTr_html, "html"); }, 2500);
            }
            if (warr_postsa_html) {
                (" Inside warr binddata");
                binddata('Warr_PostSA_Text', 'warr_postsa_text', warr_postsa_html, "html");
            }
            if (suppctr_html) {
                binddata('SupportCenter', 'supportcenter', suppctr_html, "html");
            }
            $("#getjcodes").html("Finished");
        }
 

    //Get Bill of Materials from QuoteWerks
    function getBOM() {
        var url = restURL.sql + 'services/CallStoredProcedure.rsb?Procedure=_CN_CERQBOMNewSOW&Param%3Acerq=' + model.cerq + '&@jsonp=&@authtoken=' + restURL.token;
        var html = '';
        var promise = $.ajax({
            url: url,
            contentType: "application/json;charset-uf8",
            dataType: "jsonp",
            error: function (xhr, ajaxOptions, thrownError) {
                write(xhr.status);
                write(thrownError);
                write(" An error occured while calling getBOM. ");
            }
        });
        promise.then(writeBOM);

        // Parse and assemble BOM fields.
        function writeBOM(data) {
            $.each(data.data, function (i, d) {
                if (d[4]) {
                    var partnum = d[2];
                    var jcode = d[4];
                    var desc = d[5];
                    var qty = d[7];
                    html += '<div><b>' + jcode + " | Quantity: " + qty + " | Part #: " + partnum + '</b></div><div>' + desc + '</div>'
                }
            });

            // Bind Bill of Materials to the SOW
            binddata('BOM', 'bom', html, "html");
        }  
    }
     
    // Retrieves Sales Code Solution from CRM
    function getsalescodedata() {
        $("#getPSID").html("In Progress...");
        var url = restURL.crmcustom + 'GetSalesCodeId/' + model.salescodeId + '?callback=?';
        $.getJSON(url).done(function (data) {
            var d = data.results[0];
            if (d) {
                model.layout = d.cerium_SOWLayout;
                // Get Layout to determine which Layout to us for the SOW.
                getlayoutdata(model.layout);
                model.vb = d.cerium_VerbiageBlocks;
                model.primsolname = '<div>' + d.Cerium_name + '</div>';

                // Populate variables w/ Sales Codes fields. 
                // Check for blank Sales Code verbiage. If null, then don't input anything.
                if (d.cerium_ManufacturerWarranty) {
                    model.manuwarr += '<div><b>' + d.Cerium_name + '</b></div><div>' + d.cerium_ManufacturerWarranty + '</div>';
                } else {
                    model.manuwarr += '<div><b>' + d.Cerium_name + '</b></div><div></div>';
                }
                if (d.cerium_CeriumWarranty) {
                    model.cewarranty += '<div><b>' + d.Cerium_name + '</b></div><div>' + d.cerium_CeriumWarranty + '</div>';
                } else {
                    model.cewarranty += '<div><b>' + d.Cerium_name + '</b></div><div></div>';
                }
                if (d.cerium_PrimarySOWVerbiage) {
                    model.scverb += '<div><b>' + d.Cerium_name + '</b></div><div>' + d.cerium_PrimarySOWVerbiage + '</div>';
                } else {
                    model.scverb += '<div><b>' + d.Cerium_name + '</b></div><div></div>';
                }
                
                // getsecondarySalesCode();
                $("#getPSID").html("Finished");
            } else {
                $("#getPSID").html("Error. See message below");
            }
        });
    }

    //retrieves secondary Sales Codes from CRM
    function getsecondarySalesCode() {
        $("#getSecSC").html("In Progress...");
        var url = restURL.crmcustom + "GetOppSecSalesCodeId/" + model.oppId + '?callback=?';
        $.getJSON(url).done(function (data) {
            if (data.results[0]) {
                $.each(data.results, function (i, d) {
                    model.secondarySC.push(d.cerium_salescodeid);
                });
                $("#getSecSC").html("Finished");

                //Calls func to build URL to get Secondary Sales Codes
                var buildsec = new BuildsecSCURL();
                setTimeout(function () { buildsec.buildurl(); }, 500);
            } else {

            }
        });
    }

    // Builds a URL that is used to get necessary fields for the Secondary Sales Codes (from the Sales Code table in CRM).
    function BuildsecSCURL() {
        var length = model.secondarySC.length;
        this.buildurl = function () {
            if (length > 0) {
                for (var i = 0; i < length; i++) {
                    if (i === (length - 1)) {
                        model.scId += model.secondarySC[i];
                    } else {
                        model.scId += model.secondarySC[i] + ",";
                    }
                }

                //Build the URL to lookup Secondary Sales Codes                
                model.secscurl = restURL.crmcustom + "GetSalesCodeId/" + model.scId + '?callback=?';
                this.getsecSCdata();
            }
        }

        // Call the newly built URL and retrieve Secondary Sales Code data from CRM.
        this.getsecSCdata = function () {
            var sechtml = '';
            $.getJSON(model.secscurl).done(function (data) {
                $.each(data.results, function (i, d) {

                    // Append Secondary Sales Code to previous (or Primary) Sales Code.
                    model.primsolname += '<div>' + d.Cerium_name + '</div>';

                    // Populate variables w/ Secondary Sales Codes fields. 
                    // Check for blank Secondary Sales Code verbiage. If null, do not input anything.
                    if (d.cerium_SecondarySOWVerbiage) {
                        model.scverb += '<div><b>' + d.Cerium_name + '</b></div><div>' + d.cerium_SecondarySOWVerbiage + '</div>'
                    } else {
                        model.scverb += '<div><b>' + d.Cerium_name + '</b></div><div></div>'
                    }
                    if (d.cerium_CeriumWarranty) {
                        model.cewarranty += '<div><b>' + d.Cerium_name + '</b></div><div>' + d.cerium_CeriumWarranty + '</div>';
                    } else {
                        model.cewarranty += '<div><b>' + d.Cerium_name + '</b></div><div></div>';
                    }
                    if (d.cerium_ManufacturerWarranty) {
                        model.manuwarr += '<div><b>' + d.Cerium_name + '</b></div><div>' + d.cerium_ManufacturerWarranty + '</div>';
                    } else {
                        model.manuwarr += '<div><b>' + d.Cerium_name + '</b></div><div></div>';
                    }
                })
            });
        }
    }

    // func for binding against rich text boxes then populating with data passed in
    function binddata(bindname, id, data, coerType) {
        Office.context.document.bindings.addFromNamedItemAsync(bindname, Office.BindingType.Text, { id: id }, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var binding = 'bindings#' + result.value.id;
                Office.select(binding).setDataAsync(data, { coercionType: coerType }, function (asyncResult) { });
            } else if (result.status === Office.AsyncResultStatus.Failed) {
                write("binding error for " + bindname + ". Error message: " + result.error.message);
            }
        });
    } //end bind data

    function getPMxml(type) {
        $("#getpmXML").html("In progress...");
        //points to SOW Configurator web service on port 8080
        var url = restURL.api + "api/pmlayout/getone/" + type + "?callback=?";
        $.getJSON(url).done(function (data) {
            $.each(data, function (i, d) {
                var xml = d.pmxml,
                    pos = model.pmsecpos,
                    sec = "Section" + pos,
                    id = sec.toLocaleLowerCase();
                //model.secdata.push([model.pmsecpos, xml, 6]);
                var prombind = Promisebinddata(sec, id, xml, "ooxml");
                prombind.done(function () {
                    // If there are JCodes that relate to the # of weeks of a project, bind it to the document now.
                    if (model.pmweeks) {
                        //setTimeout(function () { binddata('Projectweeks', 'projectweeks', model.pmweeks, "text"); }, 6000);
                        binddata('Projectweeks', 'projectweeks', model.pmweeks, "text");
                    }
                });  
            });
        });
        $("#getpmXML").html("Finished");
    }

    // func for binding against rich text boxes then populating with data passed in
    function Promisebinddata(bindname, id, data, coerType) {
        var def = $.Deferred();
        Office.context.document.bindings.addFromNamedItemAsync(bindname, Office.BindingType.Text, { id: id }, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var binding = 'bindings#' + result.value.id;
                Office.select(binding).setDataAsync(data, { coercionType: coerType }, function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        def.resolve();
                    }
                });
            } else if (result.status === Office.AsyncResultStatus.Failed) {
                write(" binding error for " + bindname + ". Error message: " + result.error.message);
            }
            return def.promise();
        });
        return def.promise();
    } //end bind data
               
    // Retrieve Layout from SOW Configurator
    function getlayoutdata(id) {
        $("#getseclayout").html("In Progress...");
        //points to SOW Configurator web service on port 8080
        var url = restURL.api + "api/sowlayout/getone/" + id + "?callback=?";
        $.getJSON(url).done(function (data) {
            $.each(data, function (i, d) {
                model.layoutconfig = d.sections;
            });

            // Get Section data from SOW Configurator
            getsecdata(model.layoutconfig);
            $("#getseclayout").html("Finished");
        })
        .fail(function(jqxhr, textStatus, error ) {
            var err = textStatus + ", " + error;
            write(" getlayoutdata: URL failure error: " + err);
        });
    }

    // Get Layout
    function getsecdata(layoutconfig) {
        $("#getlayoutconfig").html("In Progress...");
        var url = restURL.api + "sections?secIds=" + layoutconfig + "&callback=?",
            so = '',secxml = '',secId = '',pos = '';

        $.getJSON(url).done(function (data) {
            $.each(data, function (i, d) {
                so = d.secOrder;
                secId = d.secId;
                secxml = d.secXMLContent;
                //if Location sections exists, get locations
                if (secId === 2) {
                    getLocations(model.oppId);
                }
                //if Project Management exists, populate variable for section position to be used later in populating xml data
                if (secId === 6) {
                   model.pmsecpos = so;
                }
                //if BOM is requested, call procedure to populate array
                if (secId === 20) {
                    getBOM();
                }
                //if Configuration Overview doesn't exist, don't input PrimSolTyp and salescodeverb
                if (secId === 4) {
                    model.noSolOverview = 1;
                }
                
                // Build array for putting sections in proper order
                model.secdata.push([so, secxml, secId]);
            });

            bindtoform();
            $("#getlayoutconfig").html("Finished");
        });
    }

    // Bind the Sections to the SOW document
    function bindtoform() {
        var pos = '', sec = '', id = '';

        // Build the order that the sections need to appear on the document
        for (var i = 0; i < model.secdata.length; i++) {
            pos = model.secdata[i][0];
            model.ordsec[pos] = model.secdata[i][1];
        }

        // Append "Section" to bind to the placeholder on the SOWTemplateEmpty.doc template
        $.each(model.ordsec, function (i, d) {
            sec = "Section" + i;
            id = sec.toLocaleLowerCase();

            // Bind the XML to their respective section
            var bindx = bindxmlsecs(sec, id, d);
            bindx.then(function () {
                if (i === (model.ordsec.length - 1)) {
                    if (model.vb) {
                        getvblayout(model.vb)
                    }
                }
            });
        });
    }

    // Get verbiage blocks from SOW Configurator
    function getvblayout(vblocks) {
        var vbarr = [], newvbarr = [], blocks = [], temp = '', i=''

        //Parse field from CRM, grabbing verbiage block identifiers. Put them into an array.
        vbarr = vblocks.split(',');
        for (i = 0; i < vbarr.length; i++) {
            temp = vbarr[i].split("-VB");
            newvbarr.push(temp);
        }

        // Parse further, creating parameter for the URL for getting the XML of the verbiage blocks from SOW Configurator.
        for (i = 0; i < newvbarr.length; i++) {
            blocks.push(newvbarr[i][1]);
        }

        // Get the XML content of the verbiage blocks from SOW Configurator
        var url = restURL.api + "getvblayouts?vbIds=" + blocks + "&callback=?";
        $.getJSON(url).done(function (data) {
            $.each(data, function (i, d) {
                var vbId = d.vbId;
                var vbxml = d.vbxml;
                var sec = d.bindingname;
                var id = sec.toLocaleLowerCase();
                //setTimeout(function () { bindxmlsecs(sec, id, vbxml); }, 100);
                var bind = bindxmlsecs(sec, id, vbxml);
            });
            // Populate the rest of the data
            popdata();
            //refill Sales Code data with secondary verbiage
            fillSCdata();
        })
        .fail(function (jqxhr, textStatus, error) {
            var err = textStatus + ", " + error;
            write(" vblayouts: URL failure error: " + err);
        });
    }

    function bindxmlsecs(selection, id, xml) {
        var def = $.Deferred();
        Office.context.document.bindings.addFromNamedItemAsync(selection, Office.BindingType.Text, { id: id }, function (result) {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
               Office.select("bindings#" + id).setDataAsync(xml, { coercionType: 'ooxml' }, function (asyncResult) {
                   if (asyncResult.status == "Office.AsyncResultStatus.Failed") {
                       write(" Binding of " + selection + " failed!");
                    } else  {
                       def.resolve();
                   }
                });
                }
            return def.promise();
        }); //end bind
        return def.promise();
    }

    function getlocalxml(section, id, fileloc) {
        var output = '';
        var ooxmlreq = new XMLHttpRequest();
        ooxmlreq.open('Get', fileloc, false);
        ooxmlreq.send();

        if (ooxmlreq.status === 200) {
            output += ooxmlreq.responseText;
            binddata(section, id, output,"ooxml");
        } else if (ooxmlreq.status != 200) {
            write(" XML doesn't equal 200 ");
        }
        model.xmlengUT
    }

    function buildTOC() {
        //getlocalxml('TOC', 'toc', '../../XMLSnippets/Sections/0_TOC.xml'); //moving to 718
        setTimeout(function () { fillcontactinfo(); }, 2500);
        setTimeout(function () { fillSCdata(); }, 3000);
    }

    // remove CRLF, which otherwise come across as |r|
    function replaceCRLF(key,value) {
        var description = value.replace(/\|r\|/g, "<br>");
        var replacedtext = '<div><b>' + key + '</b></div><div>' + description + '</div></br>'
        return replacedtext;
    }


    //progress bar function: http://workshop.rs/2012/12/animated-progress-bar-in-4-lines-of-jquery/
    function progress(percent, $element) {
        var progressBarWidth = percent * $element.width() / 100;
        $element.find('div').animate({ width: progressBarWidth }, 500).html(percent + "%&nbsp;");
    }
    
    // Function that writes to a div with id='messajavge' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }

})();
