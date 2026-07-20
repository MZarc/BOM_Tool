/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: '#3B82F6',
        secondary: '#60A5FA',
        cta: '#F97316',
        background: '#0D1117',
        text: '#F8FAFC',
      },
      fontFamily: {
        mono: ['Quinngothic', 'Fira Code', 'monospace'],
        sans: ['Quinngothic', 'Fira Sans', 'sans-serif'],
      },
      backdropBlur: {
        xs: '2px',
      },
      backgroundImage: {
        'liquid-glass': 'linear-gradient(135deg, rgba(255, 255, 255, 0.1), rgba(255, 255, 255, 0.05))',
      },
      boxShadow: {
        'glass-sm': '0 8px 32px 0 rgba(31, 38, 135, 0.07)',
        'glass-md': '0 8px 32px 0 rgba(31, 38, 135, 0.15)',
        'glow-primary': '0 0 15px rgba(59, 130, 246, 0.5)',
      }
    },
  },
  plugins: [],
}
