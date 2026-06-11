import './App.css'

function App() {
  return (
    <main className="landing-page">
      <section className="login-card" aria-labelledby="page-title">
        <h1 id="page-title">Copilot Studio Insights</h1>
        <label htmlFor="environment-url">Environment URL</label>
        <input
          id="environment-url"
          name="environmentUrl"
          type="url"
          placeholder="https://your-environment.crm.dynamics.com"
        />
        <button type="button">Login</button>
      </section>
    </main>
  )
}

export default App
