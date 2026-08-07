import { useEffect, useRef, useState } from 'react'
import { motion, AnimatePresence } from 'framer-motion'

const STEPS = [
  { label: 'Initializing core engine...', progress: 15 },
  { label: 'Loading BOM processor...', progress: 38 },
  { label: 'Mounting component registry...', progress: 60 },
  { label: 'Applying design system...', progress: 80 },
  { label: 'Ready to launch ✦', progress: 100 },
]

interface Props {
  onComplete: () => void
}

export function SplashScreen({ onComplete }: Props) {
  const [stepIndex, setStepIndex] = useState(0)
  const [progress, setProgress] = useState(0)
  const [done, setDone] = useState(false)
  const rafRef = useRef<number | null>(null)
  const startRef = useRef<number | null>(null)

  const currentStep = STEPS[Math.min(stepIndex, STEPS.length - 1)]
  const targetProgress = currentStep.progress

  // Smoothly tween progress toward the current step target
  useEffect(() => {
    const animate = (ts: number) => {
      if (!startRef.current) startRef.current = ts

      setProgress(prev => {
        const diff = targetProgress - prev
        if (Math.abs(diff) < 0.3) return targetProgress
        return prev + diff * 0.07
      })

      rafRef.current = requestAnimationFrame(animate)
    }

    rafRef.current = requestAnimationFrame(animate)
    return () => {
      if (rafRef.current) cancelAnimationFrame(rafRef.current)
      startRef.current = null
    }
  }, [targetProgress])

  // Advance steps with timing
  useEffect(() => {
    const delays = [320, 620, 900, 700, 500]
    let timeout: ReturnType<typeof setTimeout>

    const advance = (index: number) => {
      if (index >= STEPS.length) return
      timeout = setTimeout(() => {
        setStepIndex(index)
        if (index === STEPS.length - 1) {
          setTimeout(() => setDone(true), 800)
        } else {
          advance(index + 1)
        }
      }, delays[index] ?? 500)
    }

    advance(0)
    return () => clearTimeout(timeout)
  }, [])

  useEffect(() => {
    if (done) {
      const t = setTimeout(onComplete, 900)
      return () => clearTimeout(t)
    }
  }, [done, onComplete])

  return (
    <AnimatePresence>
      {!done ? (
        <motion.div
          key="splash"
          initial={{ opacity: 1 }}
          exit={{ opacity: 0, scale: 1.02 }}
          transition={{ duration: 0.4, ease: 'easeOut' }}
          className="fixed inset-0 z-[9999] flex flex-col items-center justify-center bg-[#080e1a] overflow-hidden"
        >
          {/* Circuit line decorators */}
          <svg className="absolute inset-0 w-full h-full opacity-[0.04] pointer-events-none" xmlns="http://www.w3.org/2000/svg">
            <defs>
              <pattern id="grid" width="60" height="60" patternUnits="userSpaceOnUse">
                <path d="M 60 0 L 0 0 0 60" fill="none" stroke="white" strokeWidth="0.5"/>
              </pattern>
            </defs>
            <rect width="100%" height="100%" fill="url(#grid)" />
          </svg>

          {/* Core content */}
          <div className="relative flex flex-col items-center gap-10">

            {/* Logo Container */}
            <motion.div
              initial={{ opacity: 0, scale: 0.8, y: 15 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              transition={{ duration: 0.4, delay: 0.1 }}
              className="relative"
            >

              <div className="relative w-28 h-28 rounded-3xl bg-slate-900 border border-white/10 shadow-[0_24px_80px_rgba(0,0,0,0.7),inset_0_1px_0_rgba(255,255,255,0.08)] flex items-center justify-center">
                <img
                  src="/icon.png"
                  alt="BOM Tool"
                  className="w-16 h-16 object-contain brightness-125 drop-shadow-[0_0_24px_rgba(59,130,246,0.5)]"
                />
              </div>
            </motion.div>

            {/* Brand text */}
            <motion.div
              initial={{ opacity: 0, y: 14 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.55, delay: 0.45 }}
              className="flex flex-col items-center gap-2"
            >
              <h1 className="text-5xl font-black tracking-tight bg-clip-text text-transparent bg-gradient-to-br from-white via-white/95 to-white/40 select-none">
                BOM Tool
              </h1>
              <p className="text-white/25 text-xs font-black uppercase tracking-[0.35em]">
                Ultra-fast Materials Processing
              </p>
            </motion.div>

            {/* Progress Section */}
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: 0.65 }}
              className="flex flex-col gap-3 w-72"
            >
              {/* Status label */}
              <div className="flex items-center justify-between">
                <AnimatePresence mode="wait">
                  <motion.span
                    key={stepIndex}
                    initial={{ opacity: 0, y: 6 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -6 }}
                    transition={{ duration: 0.25 }}
                    className="text-[11px] font-black text-white/30 uppercase tracking-[0.15em]"
                  >
                    {currentStep.label}
                  </motion.span>
                </AnimatePresence>
                <span className="text-[11px] font-black text-primary/60 tabular-nums">
                  {Math.round(progress)}%
                </span>
              </div>

              {/* Progress bar */}
              <div className="relative h-1 w-full bg-white/5 rounded-full overflow-hidden">
                {/* Track glow */}
                <div className="absolute inset-0 rounded-full bg-gradient-to-r from-transparent via-white/5 to-transparent" />

                {/* Fill */}
                <motion.div
                  className="absolute top-0 left-0 h-full rounded-full bg-gradient-to-r from-primary via-blue-400 to-blue-300 shadow-[0_0_12px_rgba(59,130,246,0.7)]"
                  style={{ width: `${progress}%` }}
                  transition={{ duration: 0 }}
                />

                {/* Leading edge glow dot */}
                <motion.div
                  className="absolute top-1/2 -translate-y-1/2 w-3 h-3 rounded-full bg-white shadow-[0_0_16px_rgba(255,255,255,0.8)] -translate-x-1/2"
                  style={{ left: `${Math.min(progress, 99)}%` }}
                  transition={{ duration: 0 }}
                />
              </div>
            </motion.div>

            {/* Premium Luxury Developer Badge */}
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: 0.9 }}
              className="relative group cursor-default select-none mt-2"
            >
              <div className="absolute -inset-0.5 bg-gradient-to-r from-amber-500/20 via-blue-500/20 to-emerald-500/20 rounded-2xl blur-md opacity-40 group-hover:opacity-100 transition-opacity duration-500" />
              <div className="relative flex flex-col items-center justify-center px-5 py-2.5 bg-[#090f1d]/90 backdrop-blur-xl border border-white/15 rounded-2xl shadow-[0_8px_32px_rgba(0,0,0,0.6),inset_0_1px_1px_rgba(255,255,255,0.2)] group-hover:border-amber-400/40 transition-all duration-300">
                <span className="text-[8px] font-black uppercase tracking-[0.3em] text-amber-200/60 group-hover:text-amber-200/90 transition-colors leading-none mb-1">
                  DEVELOPED BY
                </span>
                <span className="text-[12px] font-black uppercase tracking-[0.2em] bg-gradient-to-r from-amber-100 via-white to-amber-300 bg-clip-text text-transparent group-hover:[text-shadow:0_0_16px_rgba(251,191,36,0.6)] transition-all leading-none">
                  MEET MISTRY
                </span>
              </div>
            </motion.div>

          </div>
        </motion.div>
      ) : null}
    </AnimatePresence>
  )
}
