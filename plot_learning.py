import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_csv("diffWheel_Learning - trials.csv")
df["correct"] = df["correct"].astype(str).str.upper() == "TRUE"
df = df.sort_values("trial_index").reset_index(drop=True)
df["trial_num"] = range(1, len(df) + 1)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4))

# Rolling accuracy across trials
ax1.plot(df["trial_num"], df["correct"].rolling(20, min_periods=5).mean(), label="rolling (20)")
ax1.axhline(0.5, color="gray", ls="--", label="chance")
ax1.set_xlabel("Trial")
ax1.set_ylabel("Accuracy")
ax1.set_ylim(0, 1)
ax1.set_title(f"Learning curve: {df['Subject_id'].iloc[0]} ({df['condition'].iloc[0]})")
ax1.legend()

# Accuracy per round
acc = df.groupby("round")["correct"].mean()
ax2.plot(acc.index, acc.values, marker="o")
ax2.axhline(0.5, color="gray", ls="--")
ax2.set_xlabel("Round")
ax2.set_ylabel("Accuracy")
ax2.set_ylim(0, 1)
ax2.set_title("Accuracy per round")

plt.tight_layout()
plt.show()
