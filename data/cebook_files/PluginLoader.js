var InstalledClient = new Object( );
InstalledClient.OC = 0;
InstalledClient.Samara = 1;
InstalledClient.AOC = 2;

//
// Name = Registered name of the Plugin
//
PluginConfigOC =
{
    IE:
    {
        Version_Name:         "CommunicatorMeetingJoinAx.JoinManager"
    },
    
    FF:
    {
        Version_CLSID:        "application/vnd.microsoft.communicator.ocsmeeting"
    }
}

PluginConfigSamara =
{
    IE:
    {
        Version_Name:         "AttendantConsoleMeetingJoinAx.JoinManager"
    },
    
    FF:
    {
        Version_CLSID:        "application/vnd.microsoft.attendantconsole.ocsmeeting"
    }
}

PluginConfigAOC =
{
    IE:
    {
        Version_Name:         "CommunicatorAttendeeMeetingJoinAx.JoinManager"
    },
    
    FF:
    {
        Version_CLSID:        "application/vnd.microsoft.communicatorattendee.ocsmeeting"
    }
}

// Format the string 
function StringFormat()
{
    var argumentList = StringFormat.arguments;
    var argsLen = argumentList.length;

    if(!argsLen)
    {
        return "";
    }
    else 
    {
        var newString = argumentList[0];
        var tmp = argsLen - 1;

        for(var i = 0; i < tmp; ++i)   
        {
            // '$' is a sepcial character in regular expression
            // if there is '$_' in string, it will break this function in IE
            // we should replace it to '$$'
            newString = newString.replace(new RegExp("%" + i, "g"), (argumentList[i + 1] + "").replace(/\$/g, "$$$$"));
        }

        return newString;
    }
}

function CreateNodeOutside(nodeType, nodeId)
{
    var node = document.createElement(nodeType || "DIV");
    document.body.appendChild(node);

    node.style.position = "absolute";
    node.style.width = "1px";
    node.style.height = "1px";
    node.style.left = "-100px";
    node.style.top = "0px";
    node.style.overflow = "hidden";
    node.style.visibility = "hidden";
    
    if (nodeId != null)
        node.id = nodeId;
    
    return node;
}

function GetBrowserTag()
{
    var browserTag = "";
    if (isIE())
    {
        browserTag = "IE";
    }
    else
    {
        // Treat all non-IE browsers the same as Firefox
        // the reason being, Opera/Chrome/Safari all look
        // for plugins in the Mozilla folder and load them
        // even if they were not installed for this browser.
        browserTag = "FF";
    }
    
    return browserTag;
}

function GetConfigForClient(installedClient, configTag)
{
    var config;
    
    switch (installedClient)
    {
        case InstalledClient.OC:
            config = PluginConfigOC[configTag];
            break;
        case InstalledClient.Samara:
            config = PluginConfigSamara[configTag];
            break;
        case InstalledClient.AOC:
            config = PluginConfigAOC[configTag];
            break;
        default:
            break;
    }

    return config;
}

//
// A general component used to load any plugin into browser (IE & Firefox)
//
function PluginLoader()
{
    this._isIE = null;
    this._name = null;
    this._clsname = null;
    this._clsid = null;
    this._initProps = null;
    
    // the loaded plugin object
    this._pluginInstance = new Object();
}

PluginLoader.IdPrefix = "_ucclient_plugin_";
PluginLoader.TagHtmlTemplateIE = "<object classid='%0'%1></object>";
PluginLoader.TagHtmlTemplateFF = "<embed type='%0'%1></embed>";

//
// Initialize it with plugin information (e.g. name, id, etc.)
//
PluginLoader.prototype.Initialize = function(name, clsname, clsid, initProps)
{
    if (this._pluginInstance.object)
    {
        return;
    }

    this._isIE = (document.all != null);
    this._name = name;
    this._clsname = clsname;
    this._clsid = clsid;
    this._initProps = initProps;
}

//
// Load plugin
//
PluginLoader.prototype.LoadPlugin = function()
{
    if (this._pluginInstance.object)
    {
        return this._pluginInstance;
    }

    this._CreatePlugin();
    return this._pluginInstance;
}

//
// UnLoad plugin
//
PluginLoader.prototype.UnloadPlugin = function()
{
    if (!this._pluginInstance.object)
    {
        return;
    }

    //if we created a dom object, remove it.
    if ((this._clsname == null) && (this._clsid != null))
    {
        try
        {
            var createdContainerNode = this._pluginInstance.object.parentNode;
            Assert(createdContainerNode != null && createdContainerNode.parentNode == document.body, "Plugin parent node is not correct");
            document.body.removeChild(createdContainerNode);
        }
        catch( ex )
        {
            // Error in removing DOM object.
            return;
        }
    }

    this._pluginInstance.object = null;
    return;
}

PluginLoader.prototype._IsPluginAvailable = function()
{
    var isAvailable = false;
    
    // Check for the availability of the Plugin
    // before we actually try to load the Plugin
    // in case of non-IE browsers: this will prevent
    // the "gold bar" indicating "Additional plugins
    // are required to display all the media on the
    // page" that comes in Firefox, etc.
    if (!this._isIE)
    {
        var mimetype = navigator.mimeTypes[this._clsid];
        
        if (mimetype)
        {
            var enabled = mimetype.enabledPlugin;
            
            if (enabled != null)
            {
                isAvailable = true;
            }
        }
    }
    
    return isAvailable;
}

//
// Create the plugin object
//
PluginLoader.prototype._CreatePlugin = function()
{
    if (this._clsname)
    {
        try
        {
            if (this._isIE)
            {
                this._pluginInstance.object = new ActiveXObject(this._clsname);
            }
            else
            {
                this._pluginInstance.object = null;
            }
        }
        catch( ex )
        {
            // Error creating ActiveX object.
        }
    }
    else if (this._clsid)
    {
        if (!this._IsPluginAvailable())
        {
            return;
        }
        
        var propStr = "";
        if (this._initProps)
        {
            var props = this._initProps;
            for (var key in props)
            {
                propStr += " " + key + "='" + props[key] + "'";
            }
        }

        var tagHtml = StringFormat(this._isIE ? PluginLoader.TagHtmlTemplateIE : PluginLoader.TagHtmlTemplateFF, this._clsid, propStr);

        try
        {
            var containerNode = CreateNodeOutside("DIV");
            containerNode.innerHTML = tagHtml;
            var node = containerNode.firstChild;
            node.id = PluginLoader.IdPrefix + this._name;
            this._pluginInstance.object = node;
        }
        catch( ex )
        {
            // Error creating DOM object.
        }
    }
}

