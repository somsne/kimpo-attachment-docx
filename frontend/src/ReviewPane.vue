<script setup lang="ts">
/**
 * 审阅栏（文档附件插件 kimpo-attachment-docx，编辑态右侧 240px）。
 *
 * 按设计稿画板：标题「批注与修订」+ 计数；修订卡片左侧蓝色竖条，作者 · 修订、时间、内容摘要，「接受 / 拒绝」；
 * 批注卡片左侧橙色竖条，作者、时间、正文，回复输入框 +「解决」（输入了回复时按钮变「回复」）。
 * 回复以缩进列在主批注下；已解决的批注折叠显示「已解决」。时间按用户时区显示（宿主 env.timezone）。
 * 新建批注：功能区「新建批注」后这里顶部出输入框，需先在文档里选中文字。
 */
import { computed, nextTick, ref, watch } from 'vue';
import type { DocxEngine, EngineChange, EngineComment, EngineSnapshot } from './engine-types';
import type { Translate } from './i18n';
import { formatLockTime } from './edit-flow';

const props = defineProps<{
  engine: DocxEngine;
  snapshot: EngineSnapshot;
  t: Translate;
  timezone: string;
  readOnly: boolean;
  /** 父组件置 true 时打开「新建批注」输入框。 */
  composing: boolean;
}>();

const emit = defineEmits<{ (e: 'update:composing', value: boolean): void }>();

const drafts = ref<Record<string, string>>({});
const newText = ref('');
const newInput = ref<HTMLTextAreaElement | null>(null);
const hint = ref('');

const roots = computed(() => props.snapshot.comments.filter(c => !c.parentId));
const repliesOf = (id: string) => props.snapshot.comments.filter(c => c.parentId === id);
const countLabel = computed(() => props.t('docx.review.count', { c: roots.value.length, t: props.snapshot.changes.length }));

function time(iso: string | null): string {
  if (!iso) return '';
  const stamp = formatLockTime(iso, props.timezone);
  // 当天只显示时分
  const today = formatLockTime(new Date().toISOString(), props.timezone).slice(0, 10);
  return stamp.startsWith(today) ? stamp.slice(11) : stamp;
}

const clip = (text: string) => (text.length > 60 ? `${text.slice(0, 60)}…` : text);

function changeText(change: EngineChange): string {
  if (change.type === 'replace') {
    return props.t('docx.review.change_replace', { deleted: clip(change.deletedText ?? ''), inserted: clip(change.insertedText ?? '') });
  }
  return props.t(`docx.review.change_${change.type}`, { text: clip(change.excerpt) });
}

function author(name: string): string {
  return name || props.t('docx.review.anonymous');
}

function submitReply(comment: EngineComment): void {
  const text = (drafts.value[comment.id] ?? '').trim();
  if (props.readOnly) return;
  if (text) {
    if (props.engine.replyComment(comment.id, text)) drafts.value[comment.id] = '';
  } else {
    props.engine.resolveComment(comment.id);
  }
}

function submitNew(): void {
  const text = newText.value.trim();
  if (!text) return;
  if (props.engine.addComment(text)) {
    newText.value = '';
    hint.value = '';
    emit('update:composing', false);
  } else {
    hint.value = props.t('docx.review.need_selection');
  }
}

watch(
  () => props.composing,
  async value => {
    if (!value) return;
    hint.value = '';
    await nextTick();
    newInput.value?.focus();
  }
);
</script>

<template>
  <aside class="kad-review" :aria-label="t('docx.review.title')" data-testid="docx-review">
    <div class="kad-review-head"><span>{{ t('docx.review.title') }}</span><span>{{ countLabel }}</span></div>

    <div v-if="composing" class="kad-card is-comment is-new" data-testid="docx-new-comment-box">
      <textarea ref="newInput" v-model="newText" class="kad-input kad-textarea" :placeholder="t('docx.review.new_placeholder')" @keydown.enter.ctrl.prevent="submitNew" />
      <p v-if="hint" class="kad-hint" role="status">{{ hint }}</p>
      <div class="kad-card-actions">
        <button type="button" class="kad-small-btn" @click="emit('update:composing', false)">{{ t('docx.review.cancel') }}</button>
        <button type="button" class="kad-small-btn is-primary" :disabled="!newText.trim()" data-testid="docx-new-comment-submit" @click="submitNew">{{ t('docx.review.add') }}</button>
      </div>
    </div>

    <div class="kad-review-list">
      <div v-for="change in snapshot.changes" :key="change.id" class="kad-card is-change" :data-change-id="change.id" data-testid="docx-change" @click="engine.focusChange(change.id)">
        <div class="kad-card-head"><span class="kad-author">{{ t('docx.review.change', { author: author(change.author) }) }}</span><span class="kad-time">{{ time(change.createdAt) }}</span></div>
        <div class="kad-card-text is-muted">{{ changeText(change) }}</div>
        <div class="kad-card-actions is-left">
          <button type="button" class="kad-small-btn" :disabled="readOnly" data-testid="docx-change-accept" @click.stop="engine.acceptChange(change.id)">{{ t('docx.cmd.accept') }}</button>
          <button type="button" class="kad-small-btn" :disabled="readOnly" data-testid="docx-change-reject" @click.stop="engine.rejectChange(change.id)">{{ t('docx.cmd.reject') }}</button>
        </div>
      </div>

      <div
        v-for="comment in roots"
        :key="comment.id"
        class="kad-card is-comment"
        :class="{ 'is-active': snapshot.activeCommentId === comment.id, 'is-resolved': comment.resolved }"
        :data-comment-id="comment.id"
        data-testid="docx-comment"
        @click="engine.focusComment(comment.id)"
      >
        <div class="kad-card-head"><span class="kad-author">{{ author(comment.author) }}</span><span class="kad-time">{{ time(comment.createdAt) }}</span></div>
        <div class="kad-card-text">{{ comment.text }}</div>
        <div v-for="reply in repliesOf(comment.id)" :key="reply.id" class="kad-reply">
          <div class="kad-card-head"><span class="kad-author">{{ author(reply.author) }}</span><span class="kad-time">{{ time(reply.createdAt) }}</span></div>
          <div class="kad-card-text">{{ reply.text }}</div>
        </div>
        <div v-if="comment.resolved" class="kad-resolved">{{ t('docx.review.resolved') }}</div>
        <div v-else class="kad-card-actions" @click.stop>
          <input v-model="drafts[comment.id]" type="text" class="kad-input kad-grow" :placeholder="t('docx.review.reply_placeholder')" :disabled="readOnly" :aria-label="t('docx.review.reply_placeholder')" @keydown.enter.prevent="submitReply(comment)" />
          <button type="button" class="kad-small-btn is-primary" :disabled="readOnly" data-testid="docx-comment-action" @click="submitReply(comment)">{{ (drafts[comment.id] ?? '').trim() ? t('docx.review.reply') : t('docx.review.resolve') }}</button>
        </div>
      </div>

      <p v-if="!snapshot.changes.length && !roots.length && !composing" class="kad-empty">{{ t('docx.review.empty') }}</p>
    </div>
  </aside>
</template>

<style scoped>
.kad-review {
  flex: none;
  width: 240px;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  min-height: 0;
  padding: 12px;
  gap: 10px;
  background: var(--kad-layout);
  border-left: 1px solid var(--kad-border);
  font-size: 12px;
}

.kad-review-head {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  color: var(--kad-text-2);
}

.kad-review-list {
  flex: 1 1 auto;
  min-height: 0;
  overflow: auto;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.kad-card {
  flex: none;
  padding: 10px;
  background: var(--kad-bg);
  border: 1px solid var(--kad-border);
  border-left: 3px solid var(--kad-warning);
  border-radius: 6px;
  cursor: pointer;
}

.kad-card.is-change {
  border-left-color: var(--kad-insert);
}

.kad-card.is-active {
  box-shadow: 0 0 0 1px var(--kad-primary);
}

.kad-card.is-resolved {
  opacity: 0.7;
}

.kad-card.is-new {
  cursor: default;
}

.kad-card-head {
  display: flex;
  justify-content: space-between;
  gap: 6px;
  margin-bottom: 4px;
}

.kad-author {
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.kad-time {
  flex: none;
  color: var(--kad-text-2);
}

.kad-card-text {
  line-height: 1.5;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.kad-card-text.is-muted {
  color: var(--kad-text-2);
}

.kad-reply {
  margin-top: 8px;
  padding-left: 8px;
  border-left: 2px solid var(--kad-border);
}

.kad-resolved {
  margin-top: 8px;
  color: var(--kad-text-2);
}

.kad-card-actions {
  display: flex;
  gap: 6px;
  margin-top: 8px;
  justify-content: flex-end;
}

.kad-card-actions.is-left {
  justify-content: flex-start;
}

.kad-small-btn {
  flex: none;
  height: 24px;
  padding: 0 8px;
  border: 1px solid var(--kad-border);
  border-radius: 4px;
  background: var(--kad-bg);
  color: var(--kad-text);
  font-size: 11px;
  cursor: pointer;
}

.kad-small-btn.is-primary {
  border-color: transparent;
  background: var(--kad-primary);
  color: #ffffff;
}

.kad-input {
  height: 24px;
  font-size: 11px;
}

.kad-textarea {
  width: 100%;
  height: 64px;
  padding: 6px 8px;
  resize: vertical;
}

.kad-hint {
  margin: 6px 0 0;
  color: var(--kad-warn-text);
}

.kad-empty {
  margin: 8px 0;
  color: var(--kad-text-2);
  text-align: center;
}
</style>
