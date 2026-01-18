// BAML client configuration for LiteLLM
// Supports any model via LiteLLM: OpenAI, Claude, Ollama, etc.

client<llm> LiteLLM {
  provider openai-generic
  options {
    base_url env.LITELLM_BASE_URL
    api_key env.LITELLM_API_KEY
    model env.LITELLM_MODEL
  }
}

// Fallback client for local development with Ollama
client<llm> Ollama {
  provider ollama
  options {
    base_url "http://localhost:11434/v1"
    model "llama3.2"
  }
}
