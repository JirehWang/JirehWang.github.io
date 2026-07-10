// Use global fetch

async function run() {
  const url = "https://script.google.com/macros/s/AKfycbxBOFeLiXu23kBMGU8iSvRyJci6fruTfk7HdahhcQFY777sCPSgasuNM7Z1CeuzuS-r/exec";
  const requestBody = {
    action: "ministry_getPageConfig",
    token: "ChurchApp-2026",
    data: {
      id: "enc_1f1d0065"
    }
  };

  console.log("Sending request to GAS...");
  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(requestBody)
  });
  
  const result = await resp.json();
  console.log("GAS Response status:", resp.status);
  console.log("GAS Response body:", JSON.stringify(result, null, 2));
}

run().catch(console.error);
