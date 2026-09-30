import { build } from "vite";
try {
  await build();
  console.log("BUILD_OK");
  process.exit(0);
} catch (e) {
  console.error("BUILD_FAIL", e && e.message ? e.message : e);
  process.exit(1);
}
