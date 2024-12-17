Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('apply-template').onclick = applyTemplate;
    }
});

const meetingTemplate = `
<h2>PURPOSE:</h2>
<h3>BACKGROUND that led to this meeting: (Attach documents or links to help if needed)</h3>
<h3>Expected ACHIEVEMENTS of the Meeting (Facilitate/discuss towards the outcome)</h3>
<ol>
    <li></li>
    <li></li>
    <li></li>
</ol>
<h3>AGENDA Items (Including estimated time, can have more than 3)</h3>
<ol>
    <li>( Min)</li>
    <li>( Min)</li>
    <li>Q&A (5 – 10 Min)</li>
</ol>
<h3>Roles of the Meeting:</h3>
<p>Facilitator:</p>
<p>Scribe:</p>
<br>
<br>
<a href="https://eur03.safelinks.protection.outlook.com/?url=https%3A%2F%2Furldefense.com%2Fv3%2F__https%3A%2F%2Fttnieurope.sharepoint.com%2F%3Ab%3A%2Fr%2Fsites%2FIntranet%2FDocuments%2FTTS-E*20Meeting-guidelines.pdf%3Fcsf%3D1%26web%3D1%26e%3Df99MeF__%3BJQ!!OS44WA!VTa8ZTwaj8PCvhTPqhFqCL40U5eJ3Rn1BchpV2s7Lf0hoWzZgV8u8Kw4d7gKGVi0dyGRO_TGA5XTSsya5PQWPiA%24&data=05%7C02%7Cit-services%40ttsystems.eu%7C667d0d5f83fe4a7d323108dd1e817bb3%7C64f8427c744041f2947aee7325623c3e%7C0%7C0%7C638700265945223886%7CUnknown%7CTWFpbGZsb3d8eyJFbXB0eU1hcGkiOnRydWUsIlYiOiIwLjAuMDAwMCIsIlAiOiJXaW4zMiIsIkFOIjoiTWFpbCIsIldUIjoyfQ%3D%3D%7C60000%7C%7C%7C&sdata=LgjB2eZFU1guB9aA9CrriUS13Ynj3mtIH%2FzKyJvvYE8%3D&reserved=0">Link to full TTS-E Meeting-guidelines</a>
`;

async function applyTemplate() {
    try {
        // First apply the template
        await setBodyTemplate();
                
        showStatus('Template and attachment added successfully!', 'success');
    } catch (error) {
        showStatus('Error: ' + error.message, 'error');
    }
}

function setBodyTemplate() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.setAsync(
            meetingTemplate,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(new Error('Failed to apply template'));
                }
            }
        );
    });
}


function showStatus(message, type) {
    const statusElement = document.getElementById('status');
    statusElement.textContent = message;
    statusElement.className = type;
    statusElement.style.display = 'block';
    setTimeout(() => {
        statusElement.style.display = 'none';
    }, 3000);
}
