import { motion } from 'framer-motion'
import { BOMDashboard } from './components/bom-dashboard'
import './index.css'

function App() {
  return (
    <motion.div
      className="min-h-screen"
      initial={{ opacity: 0, scale: 0.99 }}
      animate={{ opacity: 1, scale: 1 }}
      transition={{ duration: 0.6, ease: [0.4, 0, 0.2, 1] }}
    >
      <BOMDashboard />
    </motion.div>
  )
}

export default App
