<script setup lang="ts">
/** 线框图标（文档附件插件内部用），图形见 icons.ts。装饰性图标一律 aria-hidden，可访问名称由按钮的 aria-label 给。 */
import { computed } from 'vue';
import { ICONS, type IconName } from './icons';

const props = withDefaults(defineProps<{ name: IconName; size?: number; strokeWidth?: number }>(), { size: 16, strokeWidth: 1.8 });
const paths = computed(() => ICONS[props.name].split('|'));
</script>

<template>
  <svg :width="size" :height="size" viewBox="0 0 24 24" fill="none" stroke="currentColor" :stroke-width="strokeWidth" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">
    <path v-for="(d, i) in paths" :key="i" :d="d" />
  </svg>
</template>
