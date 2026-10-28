import 'core-js/stable'
import 'regenerator-runtime/runtime'
import ReactDOM from 'react-dom'
import Runtime, { init } from '@adobe/exc-app'
import App from './components/App'
import './index.css'

window.React = require('react')

function bootstrapRaw() {
  const mockRuntime = { on: () => {}, done: () => {} }
  const mockIms = {}
  ReactDOM.render(<App runtime={mockRuntime} ims={mockIms} />, document.getElementById('root'))
}

try {
  require('./exc-runtime')
  init(bootstrapInExcShell)
} catch (e) {
  bootstrapRaw()
}

function bootstrapInExcShell() {
  const runtime = Runtime()
  runtime.on('ready', ({ imsOrg, imsToken, imsProfile }) => {
    runtime.done()
    ReactDOM.render(
      <App runtime={runtime} ims={{ org: imsOrg, token: imsToken, profile: imsProfile }} />,
      document.getElementById('root')
    )
  })
  runtime.solution = { icon: 'AdobeExperienceCloud', title: 'Merchandising Copilot', shortTitle: 'Merch Copilot' }
  runtime.title = 'Merchandising Copilot'
}
