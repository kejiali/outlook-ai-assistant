// OpenClaw local proxy
// Cost depends on how your OpenClaw instance is configured:
//   - Local model (e.g. Ollama) → free
//   - Cloud provider (e.g. AWS Bedrock, OpenAI) → billed by that provider
var OPENCLAW_BASE_URL = "https://localhost:3000/proxy/v1";
var OPENCLAW_MODEL    = "your-openclaw-model-here"; // e.g. "amazon-bedrock/eu.anthropic.claude-sonnet-4-6"

// OpenAI (optional — uses OpenAI API credits)
var OPENAI_API_KEY = "your-openai-key-here";
