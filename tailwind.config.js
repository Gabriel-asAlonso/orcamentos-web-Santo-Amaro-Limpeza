const withOpacity = (variable) => ({ opacityValue }) => {
  if (opacityValue === undefined) {
    return `var(${variable})`;
  }

  const opacity = Number(opacityValue);
  const percentage = Number.isNaN(opacity)
    ? `calc(${opacityValue} * 100%)`
    : `${opacity * 100}%`;

  return `color-mix(in oklab, var(${variable}) ${percentage}, transparent)`;
};

module.exports = {
  darkMode: "class",
  content: [
    "./src/**/*.{html,ts}"
  ],
  theme: {
    extend: {
      fontFamily: {
        sans: ["var(--font-body)"],
        serif: ["var(--font-display)"]
      },
      colors: {
        background: withOpacity("--background"),
        foreground: withOpacity("--foreground"),
        card: withOpacity("--card"),
        "card-foreground": withOpacity("--card-foreground"),
        primary: withOpacity("--primary"),
        "primary-foreground": withOpacity("--primary-foreground"),
        secondary: withOpacity("--secondary"),
        "secondary-foreground": withOpacity("--secondary-foreground"),
        muted: withOpacity("--muted"),
        "muted-foreground": withOpacity("--muted-foreground"),
        accent: withOpacity("--accent"),
        "accent-foreground": withOpacity("--accent-foreground"),
        destructive: withOpacity("--destructive"),
        "destructive-foreground": withOpacity("--destructive-foreground"),
        border: withOpacity("--border"),
        input: withOpacity("--input"),
        ring: withOpacity("--ring"),
        "ring-offset-background": withOpacity("--background"),
        "chart-1": withOpacity("--chart-1"),
        "chart-2": withOpacity("--chart-2"),
        "chart-3": withOpacity("--chart-3"),
        "chart-4": withOpacity("--chart-4"),
        "chart-5": withOpacity("--chart-5"),
        sidebar: withOpacity("--sidebar"),
        "sidebar-foreground": withOpacity("--sidebar-foreground"),
        "sidebar-primary": withOpacity("--sidebar-primary"),
        "sidebar-primary-foreground": withOpacity("--sidebar-primary-foreground"),
        "sidebar-accent": withOpacity("--sidebar-accent"),
        "sidebar-accent-foreground": withOpacity("--sidebar-accent-foreground"),
        "sidebar-border": withOpacity("--sidebar-border"),
        "sidebar-ring": withOpacity("--sidebar-ring"),
        success: withOpacity("--success"),
        "success-foreground": withOpacity("--success-foreground"),
        warning: withOpacity("--warning"),
        "warning-foreground": withOpacity("--warning-foreground"),
        panel: withOpacity("--panel"),
        "panel-foreground": withOpacity("--panel-foreground"),
        quote: withOpacity("--quote"),
        "quote-foreground": withOpacity("--quote-foreground")
      },
      boxShadow: {
        command: "var(--shadow-command)",
        panel: "var(--shadow-panel)"
      },
      borderRadius: {
        sm: "calc(var(--radius) - 4px)",
        md: "calc(var(--radius) - 2px)",
        lg: "var(--radius)",
        xl: "calc(var(--radius) + 4px)",
        "2xl": "calc(var(--radius) + 8px)",
        "3xl": "calc(var(--radius) + 12px)",
        "4xl": "calc(var(--radius) + 16px)"
      },
      backgroundImage: {
        aurora: "var(--gradient-aurora)",
        panel: "var(--gradient-panel)"
      }
    }
  },
  plugins: []
};
