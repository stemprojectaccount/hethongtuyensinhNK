import https from "https";
https.get("https://postimg.cc/d78H7mS8", (res) => {
  let data = "";
  res.on("data", (chunk) => { data += chunk; });
  res.on("end", () => {
    const match = data.match(/<meta property="og:image" content="(.*?)"/);
    if(match) console.log(match[1]);
  });
});
