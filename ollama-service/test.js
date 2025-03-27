const { spawn } = require("child_process");

const testPrompt = "Rewrite this text to be more formal.\\n\\nThis is a test paragraph.";
console.log('Executing: ollama run llama2 "' + testPrompt + '"');

const child = spawn("ollama", ["run", "llama2", testPrompt], { shell: true });
child.stdin.end();

let outputData = "";
let errorData = "";

child.stdout.on("data", (data) => {
  outputData += data.toString();
});

child.stderr.on("data", (data) => {
  errorData += data.toString();
});

child.on("close", (code) => {
  console.log("Exit code:", code);
  console.log("Output:", outputData);
  console.log("Error:", errorData);
});
