(() => {
  const canvas = document.getElementById("board");
  const ctx = canvas.getContext("2d");
  const scoreEl = document.getElementById("score");
  const bestEl = document.getElementById("best");
  const overlay = document.getElementById("overlay");
  const overlayText = document.getElementById("overlay-text");
  const startBtn = document.getElementById("start-btn");
  const touchControls = document.getElementById("touch-controls");

  const GRID_SIZE = 20;
  const CELL = canvas.width / GRID_SIZE;
  const TICK_MS = 120;
  const BEST_KEY = "snake-best-score";

  let snake, direction, nextDirection, food, score, best, loopId, running;

  function loadBest() {
    const stored = Number(localStorage.getItem(BEST_KEY));
    return Number.isFinite(stored) ? stored : 0;
  }

  function saveBest(value) {
    localStorage.setItem(BEST_KEY, String(value));
  }

  function randomCell() {
    return {
      x: Math.floor(Math.random() * GRID_SIZE),
      y: Math.floor(Math.random() * GRID_SIZE),
    };
  }

  function placeFood() {
    let cell;
    do {
      cell = randomCell();
    } while (snake.some((s) => s.x === cell.x && s.y === cell.y));
    food = cell;
  }

  function resetState() {
    snake = [
      { x: 9, y: 10 },
      { x: 8, y: 10 },
      { x: 7, y: 10 },
    ];
    direction = { x: 1, y: 0 };
    nextDirection = direction;
    score = 0;
    scoreEl.textContent = "0";
    placeFood();
  }

  function draw() {
    ctx.fillStyle = "#001a00";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.strokeStyle = "#003300";
    ctx.lineWidth = 1;
    for (let i = 1; i < GRID_SIZE; i++) {
      ctx.beginPath();
      ctx.moveTo(i * CELL, 0);
      ctx.lineTo(i * CELL, canvas.height);
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(0, i * CELL);
      ctx.lineTo(canvas.width, i * CELL);
      ctx.stroke();
    }

    ctx.shadowBlur = 8;

    ctx.shadowColor = "#ff3b3b";
    ctx.fillStyle = "#ff3b3b";
    ctx.fillRect(food.x * CELL + 1, food.y * CELL + 1, CELL - 2, CELL - 2);

    ctx.shadowColor = "#39ff14";
    snake.forEach((seg, i) => {
      ctx.fillStyle = i === 0 ? "#a6ff8f" : "#39ff14";
      ctx.fillRect(seg.x * CELL + 1, seg.y * CELL + 1, CELL - 2, CELL - 2);
    });

    ctx.shadowBlur = 0;
  }

  function endGame() {
    running = false;
    clearInterval(loopId);
    if (score > best) {
      best = score;
      saveBest(best);
      bestEl.textContent = String(best);
    }
    overlayText.textContent = `Game over! Score: ${score}`;
    startBtn.textContent = "Play again";
    overlay.classList.remove("hidden");
  }

  function tick() {
    direction = nextDirection;
    const head = {
      x: snake[0].x + direction.x,
      y: snake[0].y + direction.y,
    };

    if (
      head.x < 0 ||
      head.x >= GRID_SIZE ||
      head.y < 0 ||
      head.y >= GRID_SIZE ||
      snake.some((seg) => seg.x === head.x && seg.y === head.y)
    ) {
      endGame();
      return;
    }

    snake.unshift(head);

    if (head.x === food.x && head.y === food.y) {
      score += 10;
      scoreEl.textContent = String(score);
      placeFood();
    } else {
      snake.pop();
    }

    draw();
  }

  function setDirection(dx, dy) {
    if (!running) return;
    if (direction.x === -dx && direction.y === -dy) return;
    nextDirection = { x: dx, y: dy };
  }

  function startGame() {
    resetState();
    running = true;
    overlay.classList.add("hidden");
    clearInterval(loopId);
    loopId = setInterval(tick, TICK_MS);
    draw();
  }

  const KEY_MAP = {
    ArrowUp: [0, -1],
    ArrowDown: [0, 1],
    ArrowLeft: [-1, 0],
    ArrowRight: [1, 0],
    w: [0, -1],
    s: [0, 1],
    a: [-1, 0],
    d: [1, 0],
  };

  window.addEventListener("keydown", (e) => {
    const mapped = KEY_MAP[e.key];
    if (mapped) {
      e.preventDefault();
      setDirection(mapped[0], mapped[1]);
    }
  });

  touchControls.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-dir]");
    if (!btn) return;
    const dirs = {
      up: [0, -1],
      down: [0, 1],
      left: [-1, 0],
      right: [1, 0],
    };
    const [dx, dy] = dirs[btn.dataset.dir];
    setDirection(dx, dy);
  });

  let touchStart = null;
  canvas.addEventListener("touchstart", (e) => {
    const t = e.touches[0];
    touchStart = { x: t.clientX, y: t.clientY };
  });
  canvas.addEventListener("touchend", (e) => {
    if (!touchStart) return;
    const t = e.changedTouches[0];
    const dx = t.clientX - touchStart.x;
    const dy = t.clientY - touchStart.y;
    if (Math.abs(dx) > Math.abs(dy)) {
      setDirection(dx > 0 ? 1 : -1, 0);
    } else {
      setDirection(0, dy > 0 ? 1 : -1);
    }
    touchStart = null;
  });

  startBtn.addEventListener("click", startGame);

  best = loadBest();
  bestEl.textContent = String(best);
  resetState();
  draw();
})();
