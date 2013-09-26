// Copyright (c) Microsoft Corporation. All rights reserved.

function Microsoft_Live_Messenger_PresenceButton_startConversation(conversationUrl)
{
    var url = conversationUrl + "&buttonhost=" + document.location.hostname;
    window.open(url, '_blank', 'height=300px,width=300px');
}

function Microsoft_Live_Messenger_PresenceButton_onStyleChange(element)
{
    if (element && element.presence)
    {
        Microsoft_Live_Messenger_PresenceButton_onPresence(element.presence);
    }
}

function Microsoft_Live_Messenger_PresenceButton_onPresence(presence)
{
    var idx = presence.id.indexOf('@');
    if (idx >= 0)
    {
        var id = presence.id.substr(0, idx);

        var element = document.getElementById('Microsoft_Live_Messenger_PresenceButton_' + id);
        if (element)
        {
            element.innerHTML = "";
        
            var conversationUrl = element.attributes['msgr:conversationUrl'].value;
            var width = parseInt(element.attributes['msgr:width'].value);
            var height = 30;

            var bgColor = element.attributes['msgr:backColor'].value;
            var altBgColor = element.attributes['msgr:altBackColor'].value;
            var color = element.attributes['msgr:foreColor'].value;

            element.className = 'Microsoft_Live_Messenger_PresenceButton';
            element.style.textAlign = 'left';
            element.style.border = 'solid 1px ' + bgColor;
            element.style.width = width + 'px';

            var outerFrame = document.createElement('div');
            outerFrame.style.border = 'solid 1px ' + altBgColor;
            outerFrame.style.backgroundColor = bgColor;
            outerFrame.style.backgroundPosition = 'center top';
            outerFrame.style.backgroundRepeat = 'repeat-x';
            outerFrame.style.overflow = 'hidden';
            outerFrame.style.textOverflow = 'ellipsis';
            outerFrame.style.whiteSpace = 'nowrap';
            outerFrame.style.width = (width - 2) + 'px';
            outerFrame.style.fontWeight = 'bold';
            outerFrame.style.height = height + 'px';

            var link = document.createElement('a');
            link.style.textDecoration = 'none';
            link.style.color = color;
            link.href = 'javascript:Microsoft_Live_Messenger_PresenceButton_startConversation("' + conversationUrl + '");';

            var innerFrame = document.createElement('div');
            innerFrame.style.padding = '4px';
            innerFrame.style.textAlign = 'center';

            var canvas = document.createElement('canvas');
            if (canvas.getContext)
            {
                // For Firefox, Safari, Opera
                canvas.width = width - 2;
                canvas.height = height - 2;
                canvas.style.position = 'absolute';
                canvas.style.zIndex = 1;

                var ctx = canvas.getContext('2d');
                var gradient = ctx.createLinearGradient(0, 0, 0, canvas.height);
                gradient.addColorStop(0, altBgColor);
                gradient.addColorStop(1, bgColor);
                ctx.fillStyle = gradient;
                ctx.fillRect(0, 0, canvas.width, canvas.height);

                innerFrame.style.position = 'absolute';
                innerFrame.style.zIndex = 2;
                innerFrame.style.paddingLeft = '0px';
                innerFrame.style.paddingRight = '0px';
                innerFrame.style.height = (height - 2) + 'px';
                innerFrame.style.width = (width - 4) + 'px';
                innerFrame.style.overflow = 'hidden';

                link.appendChild(canvas);
            }
            else
            {
                // For IE
                outerFrame.style.filter = 'progid:DXImageTransform.Microsoft.Gradient(gradientType=0,startColorStr=' + altBgColor + ',endColorStr=' + bgColor + ')';
            }

            var statusIcon = document.createElement('img');
            statusIcon.style.border = 'none';
            statusIcon.style.verticalAlign = 'middle';
            statusIcon.src = presence.icon.url;
            statusIcon.width = presence.icon.width;
            statusIcon.height = presence.icon.height;
            statusIcon.alt = presence.statusText;
            statusIcon.title = presence.statusText;

            var displayName = document.createElement('span');
            displayName.style.fontFamily = '"Segoe UI", Tahoma, Verdana, sans-serif';
            displayName.style.fontSize = '9pt';
            displayName.title = presence.displayName;

            if (displayName.innerText !== undefined)
            {
                displayName.innerText = presence.displayName;
            }
            else if (displayName.textContent !== undefined)
            {
                displayName.textContent = presence.displayName;
            }

            innerFrame.appendChild(statusIcon);
            innerFrame.appendChild(displayName);
            link.appendChild(innerFrame);
            outerFrame.appendChild(link);
            element.appendChild(outerFrame);
            
            element.presence = presence;
        }
    }
}
