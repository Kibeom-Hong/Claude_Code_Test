"""
4주차 데모: 다봉성(multimodality) — MSE 회귀 vs BeT식 이산화 vs Diffusion
numpy + matplotlib만 사용 (torch 불필요). 실행: python multimodality_demo.py  (CPU 약 1~2분)

상황: 관측 s ∈ [-1, 1] 에서, 사람 시범의 절반은 "위로 돌아가기"(a > 0), 절반은 "아래로 돌아가기"(a < 0).
      정답 행동 분포는 봉우리가 2개. 가운데(a ≈ 0)는 '장애물 정면 돌진' = 한 번도 시범에 없던 행동.
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

# 한글 폰트: 설치된 것 중 첫 번째를 사용 (Windows / Mac / Linux)
_have = {f.name for f in font_manager.fontManager.ttflist}
for _f in ["Malgun Gothic", "AppleGothic", "Apple SD Gothic Neo", "NanumGothic", "Noto Sans CJK KR", "Noto Sans CJK JP"]:
    if _f in _have:
        plt.rcParams["font.family"] = _f
        break
plt.rcParams["axes.unicode_minus"] = False

rng = np.random.default_rng(0)


# ---------------------------------------------------------------- 작은 MLP (numpy, Adam)
class MLP:
    def __init__(self, sizes, seed=0):
        r = np.random.default_rng(seed)
        self.W = [r.normal(0, np.sqrt(2 / a), (a, b)) for a, b in zip(sizes[:-1], sizes[1:])]
        self.b = [np.zeros(b) for b in sizes[1:]]
        self.m = [np.zeros_like(p) for p in self.W + self.b]
        self.v = [np.zeros_like(p) for p in self.W + self.b]
        self.t = 0

    def forward(self, x):
        self.ins, self.pres = [], []
        for i, (W, b) in enumerate(zip(self.W, self.b)):
            self.ins.append(x)
            x = x @ W + b
            if i < len(self.W) - 1:
                self.pres.append(x)
                x = x / (1 + np.exp(-x))        # SiLU
        return x

    def backward(self, g, lr=1e-3):
        gW, gb = [None] * len(self.W), [None] * len(self.W)
        for i in reversed(range(len(self.W))):
            gW[i] = self.ins[i].T @ g
            gb[i] = g.sum(0)
            if i > 0:
                g = g @ self.W[i].T
                pre = self.pres[i - 1]
                sig = 1 / (1 + np.exp(-pre))
                g = g * (sig * (1 + pre * (1 - sig)))   # d SiLU
        self.t += 1
        for k, (p, gr) in enumerate(zip(self.W + self.b, gW + gb)):
            self.m[k] = 0.9 * self.m[k] + 0.1 * gr
            self.v[k] = 0.999 * self.v[k] + 0.001 * gr ** 2
            mh = self.m[k] / (1 - 0.9 ** self.t)
            vh = self.v[k] / (1 - 0.999 ** self.t)
            p -= lr * mh / (np.sqrt(vh) + 1e-8)


# ---------------------------------------------------------------- 시범 데이터
N = 4000
s = rng.uniform(-1, 1, N)
side = rng.choice([-1.0, 1.0], N)                  # 위로/아래로 돌아가기 50:50
a = side * (0.55 + 0.25 * np.sin(2.5 * s)) + rng.normal(0, 0.04, N)
S, A = s[:, None], a[:, None]
s_test = np.linspace(-1, 1, 400)[:, None]


def minibatches(steps, bs=256):
    for _ in range(steps):
        idx = rng.integers(0, N, bs)
        yield S[idx], A[idx]


# ---------------------------------------------------------------- (1) MSE 회귀
mse = MLP([1, 64, 64, 1], seed=1)
for xs, ys in minibatches(3000):
    pred = mse.forward(xs)
    mse.backward(2 * (pred - ys) / len(xs), lr=3e-3)
a_mse = mse.forward(s_test)[:, 0]

# ---------------------------------------------------------------- (2) BeT식: k-means 이산화 + 분류 + 오프셋
K = 8
centers = np.quantile(a, np.linspace(0.05, 0.95, K))
for _ in range(30):                                # 1D k-means
    lab = np.argmin(np.abs(a[:, None] - centers[None]), 1)
    centers = np.array([a[lab == k].mean() if np.any(lab == k) else centers[k] for k in range(K)])
lab = np.argmin(np.abs(a[:, None] - centers[None]), 1)
off = a - centers[lab]
bet = MLP([1, 64, 64, 2 * K], seed=2)              # 출력: K개 bin 로짓 + K개 오프셋
for _ in range(4000):
    idx = rng.integers(0, N, 256)
    out = bet.forward(S[idx])
    logits, offs = out[:, :K], out[:, K:]
    p = np.exp(logits - logits.max(1, keepdims=True)); p /= p.sum(1, keepdims=True)
    g_logit = p.copy(); g_logit[np.arange(256), lab[idx]] -= 1        # cross-entropy
    g_off = np.zeros_like(offs)                                         # 정답 bin의 오프셋만 학습
    g_off[np.arange(256), lab[idx]] = 2 * (offs[np.arange(256), lab[idx]] - off[idx])
    bet.backward(np.concatenate([g_logit, g_off], 1) / 256, lr=3e-3)
s_rep = np.repeat(s_test, 5, 0)
out = bet.forward(s_rep)
p = np.exp(out[:, :K] - out[:, :K].max(1, keepdims=True)); p /= p.sum(1, keepdims=True)
kk = np.array([rng.choice(K, p=pi) for pi in p])
a_bet = centers[kk] + out[np.arange(len(kk)), K + kk]

# ---------------------------------------------------------------- (3) Diffusion (DDPM, T=50)
T = 50
beta = np.linspace(1e-4, 0.2, T)
abar = np.cumprod(1 - beta)


def temb(t):
    f = np.array([1, 2, 4, 8])[None]
    x = (t[:, None] / T) * f * np.pi
    return np.concatenate([np.sin(x), np.cos(x)], 1)


eps_net = MLP([1 + 1 + 8, 128, 128, 1], seed=3)
for _ in range(12000):
    idx = rng.integers(0, N, 256)
    t = rng.integers(0, T, 256)
    e = rng.normal(0, 1, (256, 1))
    at = np.sqrt(abar[t])[:, None] * A[idx] + np.sqrt(1 - abar[t])[:, None] * e
    pred = eps_net.forward(np.concatenate([at, S[idx], temb(t)], 1))
    eps_net.backward(2 * (pred - e) / 256, lr=1e-3)          # || eps - eps_theta ||^2
x = rng.normal(0, 1, (len(s_rep), 1))
for t in reversed(range(T)):
    tt = np.full(len(x), t)
    eh = eps_net.forward(np.concatenate([x, s_rep, temb(tt)], 1))
    x = (x - beta[t] / np.sqrt(1 - abar[t]) * eh) / np.sqrt(1 - beta[t])
    if t > 0:
        x += np.sqrt(beta[t]) * rng.normal(0, 1, x.shape)
a_diff = x[:, 0]

# ---------------------------------------------------------------- 그림
C_DATA, C_MSE, C_BET, C_DIF = "#9AA1AB", "#D64545", "#7C5CFF", "#F2651F"
fig, axs = plt.subplots(1, 3, figsize=(15, 4.6), sharey=True)
titles = ["(1) MSE 회귀", "(2) BeT식 이산화 (k=8)", "(3) Diffusion (DDPM)"]
for ax, ttl in zip(axs, titles):
    ax.scatter(s, a, s=3, c=C_DATA, alpha=0.35, label="사람 시범")
    ax.axhspan(-0.15, 0.15, color="#FFE3E3", zorder=0)
    ax.text(-0.98, 0.0, "장애물 (a≈0)", va="center", fontsize=10, color=C_MSE)
    ax.set_title(ttl, fontsize=14, fontweight="bold")
    ax.set_xlabel("관측 s")
    ax.spines[["top", "right"]].set_visible(False)
axs[0].set_ylabel("행동 a")
axs[0].plot(s_test[:, 0], a_mse, c=C_MSE, lw=3, label="모델 출력 (평균)")
axs[1].scatter(s_rep[:, 0], a_bet, s=4, c=C_BET, alpha=0.6, label="모델 샘플")
axs[2].scatter(s_rep[:, 0], a_diff, s=4, c=C_DIF, alpha=0.6, label="모델 샘플")
for ax in axs:
    ax.legend(loc="upper right", fontsize=9, markerscale=3)
plt.tight_layout()
plt.savefig(__file__.replace("multimodality_demo.py", "multimodality_demo.png"), dpi=160)


def in_obstacle(v):
    return float(np.mean(np.abs(v) < 0.15)) * 100


print(f"장애물 구간(|a|<0.15)에 떨어진 비율 — 시범: {in_obstacle(a):.1f}%  MSE: {in_obstacle(a_mse):.1f}%  "
      f"BeT: {in_obstacle(a_bet):.1f}%  Diffusion: {in_obstacle(a_diff):.1f}%")
